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
| `sample-issue-387.pptx`      | 手作成        | 1          | インラインテキスト装飾（太字・斜体・太字斜体）             |
| `authoring-integration.pptx` | document API  | 1          | from-scratch authoring API の package/render 統合 contract |

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
