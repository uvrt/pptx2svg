"""MicroType Express (MTX) decompression, in pure Python.

An EOT payload with ``TTEMBED_TTCOMPRESSED`` set carries its font in MicroType Express,
Monotype's font-specific compressor: an LZ77 + adaptive-Huffman coder (LZCOMP) over three
streams, whose output is a *rearranged* sfnt -- ``glyf`` is re-encoded point-by-point with
a triplet code, hinting bytecode is split out into two side streams, and ``cvt `` is delta
coded.  Decoding therefore means both decompressing and **rebuilding a TrueType file**,
which is why this module is larger than a decompressor has any right to be.

Provenance, because it matters for a format with patents on it: this is a transliteration
of libEOT (Brennan T. Vincent, MPL-2.0) -- ``src/lzcomp/*`` and ``src/ctf/*`` -- which is
in turn a port of the reference code Monotype and Microsoft submitted to the W3C in 2008
together with a royalty-free patent grant covering implementations of that submission
(https://www.w3.org/Submission/2008/01/).  The format itself is specified at
https://www.w3.org/Submission/MTX/.

**Why pure Python, and not an optional extra.**  The plan of record was a C dependency,
because LibreOffice 25.8 links libEOT for exactly this and a from-scratch MTX decoder
sounded like a project.  Two measurements changed that:

* There is no EOT decoder on PyPI at all -- neither ``libeot`` nor any binding -- so "ship
  it behind an extra like ``[metafile]``" had nothing to point at.  The extra would have
  been a package this project would first have to write and publish.
* The decode-side subset of libEOT is about 700 lines of Python, and it runs in 0.15 s
  (Inclusive Sans, 23 kB payload) to 1.06 s (Arimo Bold Italic, 164 kB payload) per face
  on a 2023 laptop.  Nine faces of the largest local test deck decode in 5.9 s total.

So the core stays standard-library-only and a bare ``pip install pptx2svg`` reads embedded
fonts.  Correctness is pinned by cross-checking against libEOT itself: all 21 payloads in
the two local template decks decode **byte-identically** to ``eot2ttf``'s output, which is
what ``tests/test_fonts_embedded.py`` asserts the structure of.

Only the decompressor is ported.  libEOT's compressor half (``Findmatch``,
``MakeCopyDecision``, ``Encode``, ``MTX_RUNLENGTHCOMP_PackData``) has no caller here.
"""

from __future__ import annotations

import struct

__all__ = ["MtxDecodeError", "decode_mtx"]

#: LZCOMP pre-loads this many bytes of synthetic data into the copy window so that bytes
#: near the start of the stream can still be expressed as copy items.  ``2*32*96 + 4*256``
#: in the reference code; the value is part of the format, not a tuning knob.
_PRELOAD_SIZE = 2 * 32 * 96 + 4 * 256  # 7168


class MtxDecodeError(Exception):
    """An MTX payload could not be decoded."""


# ---------------------------------------------------------------------------
# BITIO -- MSB-first bit reader (libEOT src/lzcomp/bitio.c).
# ---------------------------------------------------------------------------


class _BitReader:
    __slots__ = ("data", "size", "index", "count", "buffer")

    def __init__(self, data: bytes) -> None:
        self.data = data
        self.size = len(data)
        self.index = 0
        self.count = 0
        self.buffer = 0

    def bit(self) -> int:
        if self.count == 0:
            if self.index >= self.size:
                raise MtxDecodeError("bit stream exhausted")
            self.buffer = self.data[self.index]
            self.index += 1
            self.count = 8
        self.count -= 1
        return (self.buffer >> self.count) & 1

    def value(self, nbits: int) -> int:
        v = 0
        for _ in range(nbits):
            v = (v << 1) | self.bit()
        return v


# ---------------------------------------------------------------------------
# AHUFF -- adaptive Huffman, decode side (libEOT src/lzcomp/ahuff.c).
# ---------------------------------------------------------------------------


class _AdaptiveHuffman:
    """A Vitter-style adaptive Huffman tree, in parallel arrays.

    Node 1 is the root, nodes ``[1, range)`` are internal and ``[range, 2*range)`` are
    leaves at construction time; the sibling-swap in :meth:`_update_weight` then moves
    them around.  Parallel lists rather than a node class because ``read_symbol`` and
    ``_update_weight`` are the whole cost of decoding: a 320 kB face is ~300k symbols.
    """

    __slots__ = ("bio", "range", "up", "left", "right", "code", "weight", "symbol_index")

    def __init__(self, bio: _BitReader, rng: int) -> None:
        self.bio = bio
        self.range = rng
        size = 2 * rng
        self.up = [0] * size
        self.left = [0] * size
        self.right = [0] * size
        self.code = [0] * size
        self.weight = [0] * size
        self.symbol_index = [0] * rng

        for i in range(2, size):
            self.up[i] = i // 2
            self.weight[i] = 1
        for i in range(1, rng):
            self.left[i] = 2 * i
            self.right[i] = 2 * i + 1
        for i in range(rng):
            self.code[i] = -1
            self.code[rng + i] = i
            self.left[rng + i] = -1
            self.right[rng + i] = -1
            self.symbol_index[i] = rng + i

        self._init_weight(1)

        # The reference code pre-weights the model differently for the symbol coder than
        # for the two 3-bit coders.  Its test is `bitCount2 != 0`, which is set only when
        # `256 < range < 512` -- true for the symbol alphabet (256 literals + 8 lengths
        # per distance range + 3 DUP codes) and false for the length and distance
        # alphabets, which are both 8 wide.
        if 256 < rng < 512:
            self._update_weight(self.symbol_index[256])
            self._update_weight(self.symbol_index[257])
            for _ in range(12):  # DUP2
                self._update_weight(self.symbol_index[rng - 3])
            for _ in range(6):  # DUP4
                self._update_weight(self.symbol_index[rng - 2])
        else:
            for _ in range(2):
                for i in range(rng):
                    self._update_weight(self.symbol_index[i])

    def _init_weight(self, node: int) -> int:
        if self.code[node] < 0:
            self.weight[node] = self._init_weight(self.left[node]) + self._init_weight(
                self.right[node]
            )
        return self.weight[node]

    def _swap(self, a: int, b: int) -> None:
        up, left, right, code, weight = (
            self.up,
            self.left,
            self.right,
            self.code,
            self.weight,
        )
        up_a, up_b = up[a], up[b]
        up[a], up[b] = up[b], up[a]
        left[a], left[b] = left[b], left[a]
        right[a], right[b] = right[b], right[a]
        code[a], code[b] = code[b], code[a]
        weight[a], weight[b] = weight[b], weight[a]
        # The parents stay put; only the subtrees move.
        up[a] = up_a
        up[b] = up_b
        for node in (a, b):
            if code[node] < 0:
                up[left[node]] = node
                up[right[node]] = node
            else:
                self.symbol_index[code[node]] = node

    def _update_weight(self, node: int) -> None:
        weight = self.weight
        up = self.up
        while node != 1:
            current = weight[node]
            other = node - 1
            # Keep the sibling property: swap with the first node of equal weight before
            # incrementing, or the array stops being weight-ordered.
            if weight[other] == current:
                while weight[other] == current:
                    other -= 1
                other += 1
                if other > 1:
                    self._swap(node, other)
                    node = other
            weight[node] = current + 1
            node = up[node]
        weight[1] += 1

    def read_symbol(self) -> int:
        bio = self.bio
        left, right, code = self.left, self.right, self.code
        node = 1
        while True:
            # _BitReader.bit() inlined: this loop runs a few million times per face.
            count = bio.count
            if count == 0:
                if bio.index >= bio.size:
                    raise MtxDecodeError("bit stream exhausted")
                bio.buffer = bio.data[bio.index]
                bio.index += 1
                count = 8
            count -= 1
            bio.count = count
            node = right[node] if (bio.buffer >> count) & 1 else left[node]
            symbol = code[node]
            if symbol >= 0:
                break
        self._update_weight(node)
        return symbol


# ---------------------------------------------------------------------------
# LZCOMP -- decode side (libEOT src/lzcomp/lzcomp.c).
# ---------------------------------------------------------------------------

_MAX_2BYTE_DIST = 512
_LEN_WIDTH = 3
_DIST_WIDTH = 3
_BIT_RANGE = _LEN_WIDTH - 1
_LEN_MIN = 2
_DIST_MIN = 1

_RL_NORMAL = 0
_RL_SEEN_ESCAPE = 1
_RL_NEED_BYTE = 2
_RL_INITIAL = 100


def _build_preload() -> bytes:
    buf = bytearray(_PRELOAD_SIZE)
    i = 0
    for k in range(32):
        for j in range(96):
            buf[i] = k
            buf[i + 1] = j
            i += 2
    j = 0
    while i < _PRELOAD_SIZE and j < 256:
        buf[i : i + 4] = bytes((j, j, j, j))
        i += 4
        j += 1
    return bytes(buf)


_PRELOAD = _build_preload()


def _run_length_save(out: bytearray, value: int, state: int, escape: int, count: int):
    """The optional run-length layer LZCOMP wraps its output in.

    The first byte of the stream names the escape; ``escape N byte`` then means N copies
    and ``escape 0`` an escape byte.  Returns the new ``(state, escape, count)``.
    """
    if state == _RL_NORMAL:
        if value == escape:
            return _RL_SEEN_ESCAPE, escape, count
        out.append(value)
        return state, escape, count
    if state == _RL_SEEN_ESCAPE:
        if value == 0:
            out.append(escape)
            return _RL_NORMAL, escape, 0
        return _RL_NEED_BYTE, escape, value
    if state == _RL_NEED_BYTE:
        out.extend(bytes((value,)) * count)
        return _RL_NORMAL, escape, count
    return _RL_NORMAL, value, count  # _RL_INITIAL: this byte is the escape


def _lzcomp_unpack(data: bytes, version: int) -> bytes:
    bio = _BitReader(data)
    # Version 1 predates the run-length layer, so it has no flag bit to read.
    using_run_length = False if version == 1 else bool(bio.bit())

    dist_coder = _AdaptiveHuffman(bio, 1 << _DIST_WIDTH)
    len_coder = _AdaptiveHuffman(bio, 1 << _LEN_WIDTH)

    out_len = bio.value(24)

    # The symbol alphabet's width depends on how many 3-bit distance ranges the output
    # length needs, so it can only be built after reading that length.
    num_dist_ranges = 1
    while _DIST_MIN + (1 << (_DIST_WIDTH * num_dist_ranges)) - 1 < out_len:
        num_dist_ranges += 1
    dup2 = 256 + (1 << _LEN_WIDTH) * num_dist_ranges
    dup4 = dup2 + 1
    dup6 = dup4 + 1
    sym_coder = _AdaptiveHuffman(bio, dup6 + 1)

    window = bytearray(_PRELOAD_SIZE + out_len)
    window[:_PRELOAD_SIZE] = _PRELOAD
    base = _PRELOAD_SIZE

    out = bytearray()
    rl_state, rl_escape, rl_count = _RL_INITIAL, 0, 0

    read_symbol = sym_coder.read_symbol
    read_length = len_coder.read_symbol
    read_distance = dist_coder.read_symbol
    mask = 1 << _BIT_RANGE
    keep = mask - 1

    pos = 0
    while pos < out_len:
        symbol = read_symbol()
        if symbol < 256:
            value = symbol
        elif symbol == dup2:
            value = window[base + pos - 2]
        elif symbol == dup4:
            value = window[base + pos - 4]
        elif symbol == dup6:
            value = window[base + pos - 6]
        else:
            # Copy item.  The symbol carries the low bits of the length and the number of
            # 3-bit groups the distance was coded in; a continuation bit extends the
            # length across further symbols from the length coder.
            bits = symbol - 256
            ranges = (bits >> _LEN_WIDTH) + 1
            bits &= (1 << _LEN_WIDTH) - 1
            length = 0
            while True:
                more = bits & mask
                length = (length << _BIT_RANGE) | (bits & keep)
                if not more:
                    break
                bits = read_length()
            length += _LEN_MIN

            distance = 0
            for _ in range(ranges):
                distance = (distance << _DIST_WIDTH) | read_distance()
            distance += _DIST_MIN

            if distance >= _MAX_2BYTE_DIST:
                length += 1
            start = base + pos - distance - length + 1
            if start < 0:
                raise MtxDecodeError("copy item reaches behind the window")
            if pos + length > out_len:
                # The reference implementation has no such check: it writes the whole
                # copy item and only afterwards asserts `pos == out_len`, which on a
                # corrupt stream is a heap overflow rather than an error.  A valid stream
                # never reaches here, and all 21 real payloads still decode byte for byte
                # with the check in place.
                raise MtxDecodeError("copy item overruns the declared output length")
            for offset in range(length):
                value = window[start + offset]
                window[base + pos] = value
                pos += 1
                if using_run_length:
                    rl_state, rl_escape, rl_count = _run_length_save(
                        out, value, rl_state, rl_escape, rl_count
                    )
                else:
                    out.append(value)
            continue

        window[base + pos] = value
        pos += 1
        if using_run_length:
            rl_state, rl_escape, rl_count = _run_length_save(
                out, value, rl_state, rl_escape, rl_count
            )
        else:
            out.append(value)

    return bytes(out)


def _unpack_streams(data: bytes) -> list[bytes]:
    """Split an MTX payload into its three LZCOMP streams and decompress each.

    Layout: a version byte, a 24-bit copy limit, then the 24-bit start offsets of streams
    2 and 3 (stream 1 always starts at byte 10).  Stream 1 is the rearranged sfnt, 2 the
    hinting *push* data and 3 the rest of the hinting bytecode.  On the local decks
    streams 2 and 3 are 4 bytes each -- Google Fonts ship unhinted, so there is nothing
    to carry -- and stream 1 is the whole payload less ten bytes.
    """
    if len(data) < 10:
        raise MtxDecodeError("MTX payload is shorter than its header")
    version = data[0]
    offsets = (10, _u24(data, 4), _u24(data, 7))
    sizes = (
        offsets[1] - offsets[0],
        offsets[2] - offsets[1],
        len(data) - offsets[2],
    )
    streams = []
    for offset, size in zip(offsets, sizes):
        if size < 0 or offset + size > len(data):
            raise MtxDecodeError("MTX stream offsets fall outside the payload")
        streams.append(_lzcomp_unpack(data[offset : offset + size], version))
    return streams


def _u24(data: bytes, offset: int) -> int:
    return (data[offset] << 16) | (data[offset + 1] << 8) | data[offset + 2]


# ---------------------------------------------------------------------------
# CTF -- rebuild an sfnt from the three streams (libEOT src/ctf/parseCTF.c).
# ---------------------------------------------------------------------------


def _triplet_encodings() -> list[tuple[int, int, int, int, int, int, int]]:
    """``(byte_count, x_bits, y_bits, delta_x, delta_y, x_sign, y_sign)`` x 128.

    The point-delta code from https://www.w3.org/Submission/MTX/#TripletEncoding: the low
    seven bits of a point's flag byte select a row, which says how many further bytes the
    point occupies and how the x and y deltas are packed into them.
    """
    rows: list[tuple[int, int, int, int, int, int, int]] = []
    for delta in (0, 256, 512, 768, 1024):  # y-only moves
        rows.append((2, 0, 8, 0, delta, 0, -1))
        rows.append((2, 0, 8, 0, delta, 0, 1))
    for delta in (0, 256, 512, 768, 1024):  # x-only moves
        rows.append((2, 8, 0, delta, 0, -1, 0))
        rows.append((2, 8, 0, delta, 0, 1, 0))
    for delta_x in (1, 17, 33, 49):  # 4+4 bits
        for delta_y in (1, 17, 33, 49):
            for sx, sy in ((-1, -1), (1, -1), (-1, 1), (1, 1)):
                rows.append((2, 4, 4, delta_x, delta_y, sx, sy))
    for delta_x in (1, 257, 513):  # 8+8 bits
        for delta_y in (1, 257, 513):
            for sx, sy in ((-1, -1), (1, -1), (-1, 1), (1, 1)):
                rows.append((3, 8, 8, delta_x, delta_y, sx, sy))
    for sx, sy in ((-1, -1), (1, -1), (-1, 1), (1, 1)):  # 12+12 bits
        rows.append((4, 12, 12, 0, 0, sx, sy))
    for sx, sy in ((-1, -1), (1, -1), (-1, 1), (1, 1)):  # 16+16 bits
        rows.append((5, 16, 16, 0, 0, sx, sy))
    return rows


_TRIPLETS = _triplet_encodings()
assert len(_TRIPLETS) == 128, len(_TRIPLETS)


def _i16(value: int) -> int:
    value &= 0xFFFF
    return value - 0x10000 if value >= 0x8000 else value


class _In:
    """A big-endian byte reader over one decompressed CTF stream."""

    __slots__ = ("data", "pos")

    def __init__(self, data: bytes) -> None:
        self.data = data
        self.pos = 0

    def _need(self, count: int) -> None:
        if self.pos + count > len(self.data):
            raise MtxDecodeError("CTF stream is truncated")

    def u8(self) -> int:
        self._need(1)
        value = self.data[self.pos]
        self.pos += 1
        return value

    def peek(self) -> int:
        self._need(1)
        return self.data[self.pos]

    def u16(self) -> int:
        self._need(2)
        value = struct.unpack_from(">H", self.data, self.pos)[0]
        self.pos += 2
        return value

    def s16(self) -> int:
        self._need(2)
        value = struct.unpack_from(">h", self.data, self.pos)[0]
        self.pos += 2
        return value

    def u32(self) -> int:
        self._need(4)
        value = struct.unpack_from(">I", self.data, self.pos)[0]
        self.pos += 4
        return value

    def take(self, count: int) -> bytes:
        self._need(count)
        value = self.data[self.pos : self.pos + count]
        self.pos += count
        return value

    def seek(self, pos: int) -> None:
        if pos > len(self.data):
            raise MtxDecodeError("CTF stream seek past end")
        self.pos = pos

    def read255ushort(self) -> int:
        """https://www.w3.org/Submission/MTX/#id_255USHORT"""
        code = self.u8()
        if code == 253:
            return self.u16()
        if code == 255:
            return 253 + self.u8()
        if code == 254:
            return 506 + self.u8()
        return code

    def read255short(self) -> int:
        """https://www.w3.org/Submission/MTX/#id_255SHORT"""
        code = self.u8()
        if code == 253:
            return self.s16()
        sign = 1
        if code == 250:
            sign = -1
            code = self.u8()
        if code == 255:
            value = 250 + self.u8()
        elif code == 254:
            value = 500 + self.u8()
        else:
            value = code
        return value * sign


class _Out:
    """A growable big-endian sink that can seek back to patch a length or a bbox."""

    __slots__ = ("buf", "pos")

    def __init__(self) -> None:
        self.buf = bytearray()
        self.pos = 0

    def _put(self, data: bytes) -> None:
        end = self.pos + len(data)
        if end > len(self.buf):
            self.buf.extend(bytes(end - len(self.buf)))
        self.buf[self.pos : end] = data
        self.pos = end

    def u8(self, value: int) -> None:
        self._put(bytes((value & 0xFF,)))

    def u16(self, value: int) -> None:
        self._put(struct.pack(">H", value & 0xFFFF))

    def s16(self, value: int) -> None:
        self._put(struct.pack(">h", _i16(value)))

    def u32(self, value: int) -> None:
        self._put(struct.pack(">I", value & 0xFFFFFFFF))

    def raw(self, data: bytes) -> None:
        self._put(data)

    def skip(self, count: int) -> None:
        self.pos += count
        if self.pos > len(self.buf):
            self.buf.extend(bytes(self.pos - len(self.buf)))


_NPUSHB, _NPUSHW, _PUSHB, _PUSHW = 0x40, 0x41, 0xB0, 0xB8


def _push_dump(out: _Out, is_word: bool, count: int, data: list[int]) -> None:
    if count <= 0:
        return
    if count < 8:
        out.u8((_PUSHW if is_word else _PUSHB) | (count - 1))
    else:
        out.u8(_NPUSHW if is_word else _NPUSHB)
        out.u8(count)
    for value in data[len(data) - count :]:
        if is_word:
            out.s16(value)
        else:
            out.u8(value)


def _decode_push_instructions(src: _In, out: _Out, push_count: int) -> None:
    """MTX hop codes -> TrueType PUSHB/PUSHW/NPUSHB/NPUSHW.

    https://www.w3.org/Submission/MTX/#HopCodes.  Values arrive one at a time; runs of
    the same width are batched into one push instruction, which is what the encoder
    undid.
    """
    remaining = push_count
    is_word = False
    count = 0
    data: list[int] = []

    def put(value: int) -> None:
        nonlocal is_word, count
        wants_word = not (0 <= value < 256)
        if wants_word != is_word or count == 255:
            _push_dump(out, is_word, count, data)
            is_word = wants_word
            count = 0
        data.append(value)
        count += 1

    while remaining:
        code = src.peek()
        if code == 0xFB:
            # "A B 0xFB C" expands to "A B A C A".
            if remaining < 3 or len(data) < 2:
                raise MtxDecodeError("corrupt hop-code data")
            remaining -= 3
            previous = data[-2]
            src.u8()
            put(previous)
            put(src.read255short())
            put(previous)
        elif code == 0xFC:
            # "A B 0xFC C D" expands to "A B A C A D A".
            if remaining < 5 or len(data) < 2:
                raise MtxDecodeError("corrupt hop-code data")
            remaining -= 5
            previous = data[-2]
            src.u8()
            put(previous)
            put(src.read255short())
            put(previous)
            put(src.read255short())
            put(previous)
        else:
            put(src.read255short())
            remaining -= 1
    _push_dump(out, is_word, count, data)


_FLAG_ON_CURVE = 0x01
_FLAG_X_SHORT = 0x02
_FLAG_Y_SHORT = 0x04
_FLAG_X_SAME = 0x10
_FLAG_Y_SAME = 0x20


def _point_flags(x: int, y: int, on_curve: bool, first: bool) -> int:
    flags = _FLAG_ON_CURVE if on_curve else 0
    if not first and x == 0:
        flags |= _FLAG_X_SAME
    elif -256 < x < 0:
        flags |= _FLAG_X_SHORT
    elif 0 <= x < 256:
        flags |= _FLAG_X_SHORT | _FLAG_X_SAME  # X_SAME means "positive" alongside SHORT
    if not first and y == 0:
        flags |= _FLAG_Y_SAME
    elif -256 < y < 0:
        flags |= _FLAG_Y_SHORT
    elif 0 <= y < 256:
        flags |= _FLAG_Y_SHORT | _FLAG_Y_SAME
    return flags


def _read_bits(data: bytes, start: int, count: int) -> int:
    value = 0
    for i in range(count):
        bit = start + i
        value = (value << 1) | ((data[bit >> 3] >> (7 - (bit & 7))) & 1)
    return value


def _decode_simple_glyph(n_contours, streams, out, calc_bbox, x_min, y_min, x_max, y_max):
    if n_contours == 0:
        return
    src = streams[0]
    out.s16(n_contours)
    if calc_bbox:
        bbox_at = out.pos
        out.skip(8)
        x_min = y_min = 0x7FFF
        x_max = y_max = -0x8000

    else:
        out.s16(x_min)
        out.s16(y_min)
        out.s16(x_max)
        out.s16(y_max)

    total = 0
    for i in range(n_contours):
        if i == 0:
            total = 1
        total += src.read255ushort()
        out.s16(total - 1)

    flags = src.take(total)
    xs = [0] * total
    ys = [0] * total
    current_x = current_y = 0
    for i in range(total):
        byte_count, x_bits, y_bits, delta_x, delta_y, x_sign, y_sign = _TRIPLETS[
            flags[i] & 0x7F
        ]
        extra = byte_count - 1
        if x_bits + y_bits != extra * 8:
            raise MtxDecodeError("triplet encoding does not fill its bytes")
        coords = src.take(extra)
        dx = _read_bits(coords, 0, x_bits)
        dy = _read_bits(coords, x_bits, y_bits)
        xs[i] = x = _i16(x_sign * (dx + delta_x))
        ys[i] = y = _i16(y_sign * (dy + delta_y))
        # The reference accumulates in a wider type and truncates at the comparison, so
        # a font whose outline wraps int16 keeps the same (wrapped) bounding box.
        current_x = _i16(current_x + x)
        current_y = _i16(current_y + y)
        if current_x < x_min:
            x_min = current_x
        if current_x > x_max:
            x_max = current_x
        if current_y < y_min:
            y_min = current_y
        if current_y > y_max:
            y_max = current_y

    code_size_at = out.pos
    out.skip(2)
    _decode_push_instructions(streams[1], out, src.read255ushort())
    out.raw(streams[2].take(src.read255ushort()))
    code_size = out.pos - (code_size_at + 2)

    for i in range(total):
        out.u8(_point_flags(xs[i], ys[i], not flags[i] & 0x80, i == 0))
    for coords in (xs, ys):
        for i in range(total):
            value = coords[i]
            if i == 0 or value != 0:
                if -256 < value < 0:
                    value = -value
                if 0 <= value < 256:
                    out.u8(value)
                else:
                    out.s16(value)

    end = out.pos
    out.pos = code_size_at
    out.u16(code_size)
    out.pos = end
    if calc_bbox:
        out.pos = bbox_at
        out.s16(x_min)
        out.s16(y_min)
        out.s16(x_max)
        out.s16(y_max)
        out.pos = end


_COMPONENT_ARGS_WORDS = 0x0001
_COMPONENT_HAVE_SCALE = 0x0008
_COMPONENT_MORE = 0x0020
_COMPONENT_XY_SCALE = 0x0040
_COMPONENT_2_BY_2 = 0x0080
_COMPONENT_INSTRUCTIONS = 0x0100


def _decode_composite_glyph(streams, out: _Out) -> None:
    src = streams[0]
    out.s16(-1)
    for _ in range(4):  # bounding box, stored verbatim for composites
        out.s16(src.s16())
    while True:
        flags = src.u16()
        out.u16(flags)
        out.raw(src.take(2))  # glyph index
        out.raw(src.take(4 if flags & _COMPONENT_ARGS_WORDS else 2))
        if flags & _COMPONENT_2_BY_2:
            transform = 8
        elif flags & _COMPONENT_XY_SCALE:
            transform = 4
        elif flags & _COMPONENT_HAVE_SCALE:
            transform = 2
        else:
            transform = 0
        out.raw(src.take(transform))
        if not flags & _COMPONENT_MORE:
            break
    if flags & _COMPONENT_INSTRUCTIONS:
        count_at = out.pos
        out.skip(2)
        _decode_push_instructions(streams[1], out, src.read255ushort())
        out.raw(streams[2].take(src.read255ushort()))
        count = out.pos - (count_at + 2)
        if count > 0:
            end = out.pos
            out.pos = count_at
            out.u16(count)
            out.pos = end


def _decode_glyph(streams, out: _Out) -> None:
    src = streams[0]
    n_contours = src.s16()
    if n_contours < 0:
        _decode_composite_glyph(streams, out)
        return
    x_min = y_min = x_max = y_max = 0
    calc_bbox = True
    if n_contours == 0x7FFF:
        # The sentinel means the encoder could not derive the box from the points, so it
        # stored the real count and the real box instead.
        n_contours = src.s16()
        x_min, y_min, x_max, y_max = src.s16(), src.s16(), src.s16(), src.s16()
        calc_bbox = False
    _decode_simple_glyph(n_contours, streams, out, calc_bbox, x_min, y_min, x_max, y_max)


def _unpack_cvt(src: _In, offset: int) -> bytes:
    """``cvt `` is stored as deltas in a variable-length code; undo both."""
    src.seek(offset)
    count = src.u16()
    out = _Out()
    value = 0
    for _ in range(count):
        code = src.u8()
        if code >= 248:
            delta = 238 * (code - 247) + src.u8()
        elif code >= 239:
            delta = -(238 * (code - 239) + src.u8())
        elif code == 238:
            delta = src.s16()
        else:
            delta = code
        value = _i16(value + delta)
        out.s16(value)
    return bytes(out.buf)


def _parse_ctf(streams: list[bytes]) -> bytes:
    main = _In(streams[0])
    ins = (main, _In(streams[1]), _In(streams[2]))

    main.u32()  # scaler type; the output is always written as 0x00010000
    num_tables = main.u16()
    main.u16()  # searchRange
    main.u16()  # entrySelector
    main.u16()  # rangeShift

    tables: list[dict] = []
    for _ in range(num_tables):
        tag = main.take(4)
        if tag in (b"hdmx", b"VDMX"):
            # Device-metrics tables are keyed to the *original* glyph set and libEOT does
            # not rebuild them; carrying them through unchanged would be worse than
            # dropping them, since a rasteriser that trusts hdmx would then disagree with
            # hmtx.  Neither table affects the advance widths we measure from.
            main.pos += 12
            continue
        main.pos += 4  # checksum; recomputed on output
        tables.append(
            {"tag": tag, "offset": main.u32(), "size": main.u32(), "buf": None}
        )

    by_tag = {table["tag"]: table for table in tables}
    glyf = by_tag.get(b"glyf")
    loca = by_tag.get(b"loca")
    head = by_tag.get(b"head")
    maxp = by_tag.get(b"maxp")

    for table in tables:
        tag = table["tag"]
        if tag in (b"loca", b"glyf"):
            continue  # rebuilt below
        if tag == b"cvt ":
            table["buf"] = _unpack_cvt(main, table["offset"])
            continue
        main.seek(table["offset"])
        buf = main.take(table["size"])
        if tag == b"head":
            if len(buf) < 54:
                raise MtxDecodeError("malformed head table")
            buf = bytearray(buf)
            buf[8:12] = b"\0\0\0\0"  # checkSumAdjustment, recomputed in _dump_sfnt
            buf = bytes(buf)
        table["buf"] = buf

    if head is None:
        raise MtxDecodeError("CTF font has no head table")
    if maxp is None:
        raise MtxDecodeError("CTF font has no maxp table")
    if b"hmtx" not in by_tag:
        raise MtxDecodeError("CTF font has no hmtx table")
    if len(head["buf"] or b"") < 54 or len(maxp["buf"] or b"") < 6:
        raise MtxDecodeError("CTF head or maxp table is too short to read")

    if glyf is not None:
        if loca is None:
            loca = {"tag": b"loca", "offset": 0, "size": 0, "buf": None}
            tables.append(loca)
        index_to_loc_format = struct.unpack_from(">h", head["buf"], 50)[0]
        num_glyphs = struct.unpack_from(">H", maxp["buf"], 4)[0]
        main.seek(glyf["offset"])
        ins[1].seek(0)
        ins[2].seek(0)
        glyf_out = _Out()
        loca_out = _Out()
        short_loca = index_to_loc_format == 0
        if short_loca:
            loca_out.u16(0)
        else:
            loca_out.u32(0)
        for _ in range(num_glyphs):
            _decode_glyph(ins, glyf_out)
            if glyf_out.pos % 2:
                glyf_out.u8(0)
            if short_loca:
                loca_out.u16(glyf_out.pos // 2)
            else:
                loca_out.u32(glyf_out.pos)
        glyf["buf"] = bytes(glyf_out.buf)
        loca["buf"] = bytes(loca_out.buf)

    return _dump_sfnt(tables)


def _dump_sfnt(tables: list[dict]) -> bytes:
    """Write the tables back out as a TrueType file, with fresh checksums."""
    out = _Out()
    count = len(tables)
    largest_power = 1
    remaining = count
    entry_selector = 0
    while remaining > 1:
        largest_power *= 2
        entry_selector += 1
        remaining //= 2
    search_range = largest_power * 16

    out.u32(0x00010000)
    out.u16(count)
    out.u16(search_range)
    out.u16(entry_selector)
    out.u16(count * 16 - search_range)

    directory_at = out.pos
    out.skip(16 * count)

    total = 0
    head_offset = None
    for table in tables:
        table["out_offset"] = out.pos
        if table["tag"] == b"head":
            head_offset = out.pos
        buf = table["buf"] or b""
        padded = buf + bytes(-len(buf) % 4)
        checksum = 0
        for i in range(0, len(padded), 4):
            checksum = (checksum + struct.unpack_from(">I", padded, i)[0]) & 0xFFFFFFFF
        table["checksum"] = checksum
        table["out_size"] = len(buf)
        out.raw(padded)
        total = (total + checksum) & 0xFFFFFFFF

    if head_offset is None:
        raise MtxDecodeError("CTF font has no head table")

    out.pos = directory_at
    for table in tables:
        out.raw(table["tag"])
        out.u32(table["checksum"])
        out.u32(table["out_offset"])
        out.u32(table["out_size"])

    prefix = bytes(out.buf[: out.pos])
    for i in range(0, len(prefix), 4):
        total = (total + struct.unpack_from(">I", prefix, i)[0]) & 0xFFFFFFFF

    # 0xB1B0AFBA is the constant the TrueType spec defines for checkSumAdjustment.
    out.pos = head_offset + 8
    out.u32((0xB1B0AFBA - total) & 0xFFFFFFFF)
    return bytes(out.buf)


def decode_mtx(data: bytes) -> bytes:
    """Decompress an MTX payload into a TrueType file."""
    return _parse_ctf(_unpack_streams(data))
