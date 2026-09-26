"""Order-preserving process pools for the development tools: ``tools/fidelity.py``,
``tools/pdf_svg.py`` and ``tools/extract_font_metrics.py``.

Standard library only, and nothing here is specific to PowerPoint, so the sibling
``docx2svg`` can take the file as it is.

The contract every caller relies on: :func:`pool_map` returns exactly what
``list(map(function, tasks))`` would, in the same order, and with ``jobs == 1`` it *is*
that -- no pool, no pickling, no other process.  Each task is computed whole in one
process, so no floating-point reduction is ever split or reordered across workers, and a
parallel run's numbers are the serial run's bit for bit.
"""

from __future__ import annotations

import os


def physical_memory() -> int | None:
    """Bytes of RAM, or ``None`` where the platform will not say (Windows)."""
    try:
        return os.sysconf("SC_PAGE_SIZE") * os.sysconf("SC_PHYS_PAGES")
    except (AttributeError, ValueError, OSError):
        return None


def default_jobs(worker_memory: int = 0) -> int:
    """Every logical core, unless the machine's memory holds fewer workers of
    ``worker_memory`` bytes each than that."""
    jobs = os.cpu_count() or 1
    memory = physical_memory()
    if memory and worker_memory:
        jobs = min(jobs, max(1, memory // worker_memory))
    return jobs


def pool_map(function, tasks, jobs: int, initializer=None, initargs=()) -> list:
    """``list(map(function, tasks))`` over ``jobs`` processes, in the order of ``tasks``.

    ``jobs == 1`` (or a single task) runs in this process, ``initializer(*initargs)``
    first: the serial path exactly.  Otherwise a pool of fresh interpreters -- ``spawn``
    on every platform, since a fork would copy PyMuPDF's and resvg's native state
    half-initialised -- runs ``initializer`` once each, then the tasks one at a time, so
    one long task does not hold a queue of short ones behind it.  ``function`` and the
    tasks must pickle: a module-level function, plain data.
    """
    tasks = list(tasks)
    if jobs <= 1 or len(tasks) <= 1:
        if initializer is not None:
            initializer(*initargs)
        return list(map(function, tasks))
    import multiprocessing
    from concurrent.futures import ProcessPoolExecutor

    with ProcessPoolExecutor(max_workers=min(jobs, len(tasks)), mp_context=multiprocessing.get_context("spawn"),
                             initializer=initializer, initargs=initargs) as pool:
        return list(pool.map(function, tasks, chunksize=1))
