# shared-fixtures

実アプリケーション（PowerPoint / Google Slides / md-pptx など）が生成した実 PPTX ファイルと、公開 authoring API の固定 integration baseline を格納するディレクトリ。

## 用途

このディレクトリのファイルは複数のテストスイートから共有される:

| テスト               | ファイル                          | 目的                                              |
| -------------------- | --------------------------------- | ------------------------------------------------- |
| E2E スモークテスト   | `e2e/smoke.test.ts`               | エラーなく変換できること・主要要素の SVG 出力確認 |
| スナップショット VRT | `vrt/snapshot/regression.test.ts` | レンダリング結果のリグレッション検出              |

## プログラム合成フィクスチャとの違い

`vrt/snapshot/fixtures/` にあるフィクスチャは `create-fixtures.ts` でプログラム的に生成される。
このディレクトリのファイルは **手動で作成・管理される実 PPTX** であり、プログラム合成では再現しにくい以下の構造を含む:

- テーマフォント参照 (`+mj-lt`, `+mn-lt`)
- `presentation.xml` の `defaultTextStyle`
- スライドマスターの `txStyles`
- スタイル参照 (`sp.style` の `lnRef` / `fillRef` / `effectRef`)
- PowerPoint バージョン固有の XML 構造

## ファイル一覧

| ファイル                     | 作成元        | スライド数 | 主な内容                                                   |
| ---------------------------- | ------------- | ---------- | ---------------------------------------------------------- |
| `real-basic-theme.pptx`      | Google Slides | 2          | タイトル・コンテンツ・テーブル・画像・テーマフォント参照   |
| `real-product-page.pptx`     | 手作成        | 1          | 角丸矩形・楕円・テキストボックス                           |
| `real-financial-report.pptx` | 手作成        | 4          | チャート（棒グラフ・円グラフ等）・テキスト                 |
| `sample.pptx`                | md-pptx 生成  | 6          | 日本語テキスト・箇条書き・テキスト装飾                     |
| `sample-cjk.pptx`            | 本リポジトリの `tools/make_cjk_deck.py` が `sample.pptx` から**派生生成** | 6 | 同上（テーマの東アジア系フェイスのみ差し替え）             |
| `sample-issue-387.pptx`      | 手作成        | 1          | インラインテキスト装飾（太字・斜体・太字斜体）             |
| `authoring-integration.pptx` | document API  | 1          | from-scratch authoring API の package/render 統合 contract |
| `chart-gallery.pptx`         | 本リポジトリの `tools/make_chart_gallery.py` が**生成** | 17 | チャート種別を 1 スライド 1 種で網羅 |
| `feature-sweep.pptx`         | 本リポジトリの `tools/make_feature_sweep.py` が**生成** | 13 | コーパスのどのデッキも踏まない機能を 1 スライド 1 機能で網羅 |

## `sample-cjk.pptx`（`sample.pptx` からの派生。**取得したままのファイルではない**）

`tools/make_cjk_deck.py` が `sample.pptx` から生成する。変更点は**テーマだけ**で、
theme1 / theme2 の major・minor 両フォントコレクションについて

- `<a:ea typeface=""/>` → `<a:ea typeface="Noto Sans JP"/>`
- `<a:font script="Jpan" typeface="ＭＳ Ｐゴシック"/>` → 同じく `Noto Sans JP`

の 8 箇所を書き換える。スライド・レイアウト・マスター・本文テキストは 1 文字も
変更していない。したがって**本文の内容は md-pptx が生成したそのまま**だが、
**フォント指定は原本のものではない**。

存在理由は `tools/fidelity.py` で**採点できる日本語デッキが 1 つもなかった**こと。
`sample.pptx` が指定する ＭＳ Ｐゴシック は、このマシンの PowerPoint では綴りを変えても
解決できず（`tools/make_cjk_deck.py` の docstring に 7 回のエクスポート実測表がある）、
Microsoft のフォントなのでこちらで導入する権利もない。つまり**原本のフェイスは
どちらのレンダラも描かない**ので、比較しても測っているのはフォント解決であって
レイアウトではない。Noto Sans JP は本リポジトリが同梱し、`~/Library/Fonts` に入れれば
PowerPoint も描く唯一の日本語フェイスであり、そのとき初めて両者が同じ字形を描く。

`sample.pptx` 自体は**一切変更していない**。空の `<a:ea>` は実在するケースであり
（`src/pptx2svg/text/fontmap.py`）、既存の計測結果の入力でもあるため、
書き換えるのではなく `tests/deckbuilder.py` と同じく**派生**させている。

## `real-college-template.pptx`（**リポジトリには含めない**）

Dickinson College が公開しているサンプルプレゼンテーション。**第三者の著作物であり、
再配布する権利はこちらにない**ため、`tests/fixtures/` には置かず、gitignore された
`scratch/` に各自で配置する。必要とするテストは、無ければ skip する
（`tests/conftest.py` の `college_template` fixture が唯一の参照点）。

| 項目           | 値                                                                             |
| -------------- | ------------------------------------------------------------------------------ |
| 取得元         | <https://www.dickinson.edu/download/downloads/id/1076/sample_powerpoint_slides.pptx> |
| 取得日         | 2026-09-11                                                                     |
| sha256         | `ac7f2627645042190df3244cc25929f4b006d144fc2cac520e79ab376197bbbf`             |
| サイズ         | 647,237 bytes                                                                  |
| 作成アプリ     | Microsoft Office PowerPoint 12.0（Office 2007）、2012-03-02 作成                |
| テーマ         | `Dickinson_Template_red`（accent1 = `#C00000`、dk1 = `#151515`）               |

内容:

- 9 スライド。4:3（On-screen Show）。
- `ppt/media/` に EMF 2 点（いずれも埋め込み PDF プレビュー付き）、PNG 1 点、JPEG 1 点。
- `ppt/charts/chart1.xml` は積み上げ縦棒チャート 1 点。`ppt/embeddings/` の `.xlsx` を
  `c:externalData` として参照し、`c:userShapes`（`ppt/drawings/drawing1.xml`）も持つ。
- スライド 5 はデッキ自身の `tableStyles.xml` が定義する
  "Medium Style 2 - Accent 1" を使うテーブル。

**このデッキがコーパスに入った理由**は、要求フェイス（Arial / Calibri / Wingdings）が
すべてこのマシンに存在し、実 PPTX として初めて **スキップされずに全スライドを採点できる**
ためである。他の実 PPTX はいずれもフォント差で `tests/fidelity-baselines.json` 上
`skipped` になっている。

このデッキだけが持っていて、他のコーパスが持っていなかったもの:

- **負の値を持つチャート**。2002 カテゴリが `-1.0` で、系列は `c:invertIfNegative` を
  書いていない。`src/pptx2svg/resolve/chart.py` の既定値がここで誤りだと判明した。
- **`c:idx` と `c:order` が食い違う系列**。`(idx 2, order 0)` と `(idx 0, order 1)` で、
  アクセント色の周回が `c:idx` 基準であることを初めて区別できた。
- **マスター / レイアウトの `b="1"` だけで太字になるタイトル**。
- **`a:spcBef` と `a:spcAft` が加算される**ことを示す段落間隔。

なお 647KB は下記「1ファイルあたり 500KB 以下」の目安を超える。内訳は
`ppt/media/image4.jpeg` の 441KB がほぼすべてである。いずれにせよリポジトリには
含めないため、この目安には抵触しない。

上記 4 点の知見は、いずれもこのデッキを必要としないテストとして
`tests/test_chart.py` / `tests/test_render.py` に定着させてある。
デッキ本体が無くても回帰は検出できる。

## `feature-sweep.pptx`（`tools/make_feature_sweep.py` が**生成**。取得ファイルではない）

### 存在理由

このプロジェクトは「パーサは読んでいるのにレンダラが捨てている」という同じ欠陥に
3 回刺されている（`flat_chart_kind` の 3-D 綴り消失、ChartEx フレームの誤診断、
`a:clrChange`）。3 つ目はコードを読んで見つかったのではなく**カバレッジテスト**で
見つかった。理由は単純で、**コーパスのどのデッキもその機能を踏んでいない**からである。
動くスナップショットも下がる fidelity スコアも存在しないので、壊れていても何も鳴らない。

このデッキはその穴を埋める。`chart-gallery.pptx` と同じ方針で **1 スライド 1 機能**に
してあるので、スコアが原因まで一意に落ちる。

### スナップショットを「正しさの主張」として読まないこと

13 スライドのうち**正しく描けているのは 7 枚だけ**である。残りは*意図的に直していない*
機能を、現状のまま記録するために置いてある。どれがどちらかは下表と
`tools/make_feature_sweep.py` の `SLIDES` に書いてある。

| # | 機能 | 状態 |
| --- | --- | --- |
| 1 | `a:clrChange` | 描画する（本スイープで実装） |
| 2 | `asvg:svgBlip` | 描画する（本スイープで実装） |
| 3 | `mc:AlternateContent` の分岐選択 | 描画する（本スイープで修正） |
| 4 | `a:buBlip` | 描画する（本スイープで実装） |
| 5 | `a:buSzPts` | 描画する（本スイープで実装） |
| 6 | `a:alphaModFix` | 描画する（本スイープで実装） |
| 7 | `a:ln@cmpd` | **描かない**。1 本の線に潰し `line-compound-flattened` で申告 |
| 8 | `a:lnL/R/T/B@cmpd`（表の罫線） | **描かない**。同上、警告はスライドごとに 1 回 |
| 9 | `a:path@path`（`circle` / `rect` / `shape`） | **一部のみ**。3 種とも同じ radial になる |
| 10 | `a:tile@flip` / `@algn` | **描かない**。加えて**タイル寸法が約 8.3 倍**（下記） |
| 11 | `a:bodyPr@anchorCtr` | **描かない**。`src/` に一度も現れない |
| 12 | `a:outerShdw@rotWithShape` / `a:blur@grow` | **描かない** |
| 13 | `a:pattFill` | 描画するが**タイルが細かすぎる**（下記） |

### 採点できるように作ってある

`tools/fidelity.py` がスキップするデッキは価値が大きく下がる（コーパス 7 枚のうち 3 枚は
PowerPoint 側がフォントを置換するためスキップされている）。そこで `make_chart_gallery.py`
と同じ制約を課してある。

- **CJK を 1 文字も使わない**。Latin-1 の外の文字も使わない。
- **書体を一切名指ししない**。`a:latin` / `a:ea` / `a:cs` に加え、`a:buFont` も書かない
  （`fidelity.requested_faces` はこれを数える）。全スライドがテーマの Aptos だけを使う。

結果として `requested_faces` は `Aptos` と `Aptos Display` の 2 つだけを返し、
このマシンの PowerPoint は両方をネイティブに描くので**スキップされず採点される**
（`fonts 2/2`）。

### 出自

すべて本リポジトリのもの。package skeleton（theme / master / layout / `presentation.xml`）は
`authoring-integration.pptx`（本リポジトリの authoring API が生成したもの）から取り、
master と layout は空にしてある。スライド 13 枚と `ppt/media/` の画像 4 点はすべて
`tools/make_feature_sweep.py` が書く。ラスタは `zlib` で数十バイトずつ符号化し、
ベクタは矩形の並びである。第三者のデッキも第三者の素材も含まない。

画像を「厳密な単色ブロック」にしてあるのは意図的である。`a:clrChange` は**完全一致**で
しか置換しないので写真では鍵にする色が存在せず、PNG は可逆なので PowerPoint が復号する
色と resvg が復号する色が一致する。

### このデッキが見つけた欠陥

作った時点で 2 つ出た。どちらもコーパスが一度も踏んでいなかったので、それまで不可視だった。

- **タイル画像塗りが一切タイルしていなかった**（スライド 10）。`<pattern>` に `viewBox` が
  無く、中の `<image width="100%">` が**タイルではなくビューポート**（960px のスライド全体）
  に対して解決されていたため、各タイルには画像の隅が極端に拡大されたものだけが入り、
  結果は単色の塗りつぶしだった。`viewBox="0 0 1 1"` を与えて修正済み。
  コーパスに `a:tile` は 1 つも無い。
- **タイル寸法の解釈が誤っている**（スライド 10、未修正）。`a:tile@sx` を
  「shape の bounding box に対する割合」として扱っているが、OOXML では
  **画像の原寸に対する倍率**である。PowerPoint の書き出しを実測すると 136.8 pt 幅の箱で
  周期 **9.900 pt**、こちらのモデルは `0.6 × 136.8 = 82.08 pt` で **約 8.3 倍**大きい。
  ただし素直な予測（32 px を 96 dpi として 24 pt × 0.6 = 14.4 pt）とも一致しないため、
  正しい規則は `sx` / 画像画素数 / 箱の寸法を振ったプローブ実測が要る。本スイープの範囲外。
- **`a:pattFill` のタイルが PowerPoint より細かい**（スライド 13、未修正）。`horz` で
  PowerPoint の約 2 倍の本数が出る。`dkDnDiag` は密度だけでなく見た目も異なる
  （PowerPoint は太い斜帯、こちらは細かい網）。これも実測が要る。

### 再生成と再エクスポート

```bash
python3 tools/make_feature_sweep.py                        # フィクスチャを書き直す
python3 tools/make_feature_sweep.py ~/pptx2svg-oracle/feature-sweep.pptx
osascript tools/powerpoint_export_pdf.applescript \
    ~/pptx2svg-oracle/feature-sweep.pptx ~/pptx2svg-oracle/feature-sweep.pdf
python3 -m pytest tests/test_vrt.py --update-snapshots
python3 tools/fidelity.py --update
```

**受け入れ条件は「PowerPoint が開いて書き出せること」**である。修復（`[Repaired]`）される
デッキはフィクスチャではない。`tests/deckbuilder.py` も `make_feature_sweep.py` も
content-type の `Override` に先頭スラッシュを**自分で付ける**ので、`f"/{part}"` を渡すと
`PartName="//ppt/..."` になり修復される。


## Authoring integration fixture の再生成

`authoring-integration.pptx` は `@pptx-glimpse/document` の package root から公開された
`createPptx` と authoring API だけで生成する。次のコマンドで再生成できる。

```bash
npm run fixture:authoring
npm run test -- e2e/authoring-integration.test.ts
```

baseline は、公開 authoring contract または生成される汎用 PPTX / OOXML package structure を
意図的に変更した場合だけ更新する。ZIP metadata や byte-for-byte equality は baseline contract に
含めない。更新時は integration test で package relationship、content type、part path、ID、reader / writer
round-trip、core document-path renderability が引き続き成立することを確認する。

## ファイルを追加する場合

1. `shared-fixtures/` に PPTX ファイルを配置する
2. `e2e/smoke.test.ts` にスモークテストを追加する
3. `vrt/snapshot/vrt-cases.ts` の `SHARED_FIXTURE_CASES` にエントリを追加する
4. `npm run vrt:snapshot:update` でスナップショットを生成する

ファイルサイズはリポジトリサイズへの影響を抑えるため **1ファイルあたり 500KB 以下** を目安とする。
ライセンス上問題のない PPTX のみ追加すること（自作または OSSテンプレート）。
