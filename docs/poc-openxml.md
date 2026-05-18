# POC 記録 → 本実装候補化記録: Markdig + Open XML SDK 方式

v0.5.1 で POC 実装を行い、v0.5.2 で本実装候補として機能拡張し、v0.5.3 で UI から選択可能にした記録。
`docs/engine-comparison.md` の推奨を受けて、実用可能かを検証・整備する。

---

## 1. 実装概要

### 対象ファイル

| ファイル | 説明 |
|----------|------|
| `src/md2doc/OpenXmlConverter.cs` | 変換クラス（Word COM 不要） |
| `tests/Md2Doc.Tests/OpenXmlConverterTests.cs` | テスト（全 Word 不要、26 テスト） |
| `tests/Md2Doc.Tests/DocxInspector.cs` | docx XML 検査ヘルパー（v0.5.2 で機能拡張） |

### 依存ライブラリ

| パッケージ | バージョン | ライセンス |
|-----------|-----------|-----------|
| `Markdig` | 0.37.0 | BSD-2-Clause |
| `DocumentFormat.OpenXml` | 3.2.0 | MIT |

---

## 2. 実装アーキテクチャ

```
Markdown 文字列
    ↓ Markdig.Parse()
Markdig AST（MarkdownDocument）
    ↓ OpenXmlConverter.AppendBlock()（switch 振り分け）
OOXML 要素（Paragraph / Run / Table / Break 等）
    ↓ WordprocessingDocument.Save()
.docx ファイル
```

### Word COM 方式との比較

| 処理 | Word COM 方式 | Open XML 方式 |
|------|--------------|--------------|
| Markdown 解析 | 手製 `ParseBlocks()` | Markdig AST（外部ライブラリ） |
| 段落生成 | `doc.Content.Text = 全文` | `body.Append(new Paragraph(...))` |
| スタイル適用 | `para.Range.Style = -N` | `new ParagraphStyleId { Val = "HeadingN" }` |
| 箇条書き | `para.Range.Style = -47` + 2 パス伝播除去 | NumberingInstance（確定的、伝播なし） |
| 順序付きリスト | 未対応 | NumberingInstance（LevelOverride で独立番号） |
| 表 | `doc.Tables.Add()` | `new Table(...)` |
| 改ページ | `doc.Content.Text` 内に `\f` | `new Break { Type = BreakValues.Page }` |
| インライン書式 | 除去のみ | `RunProperties` に `Bold` / `Italic` / 等幅フォント |
| ヘッダー | `doc.Sections(1).Headers(1)` | `HeaderPart` + `SectionProperties.HeaderReference` |
| フッター（ページ番号） | `doc.Sections(1).Footers(1)` | `FooterPart` + `FieldChar / FieldCode(" PAGE ")` |
| 水平線 | `para.Range.Borders(-3)` | `ParagraphBorders.BottomBorder` |
| ソフトリターン | `para.Range.InsertAfter(wdSoftReturn)` | `new Break()` |
| COM 安定化 | 4 フェーズ + 2 パス補正 | 不要（直接 XML 構築） |

---

## 3. 対応要素（v0.5.2 時点）

| Markdown 要素 | 対応状況 | 備考 |
|--------------|---------|------|
| 見出し H1〜H3 | ✅ | スタイル定義付き（16/14/12pt 太字） |
| 通常段落 | ✅ | |
| 箇条書き（`-` / `*`） | ✅ | `•` 記号、Word リスト形式 |
| 順序付きリスト（`1.`） | ✅ | `%1.` 形式、複数リスト独立番号 |
| 表 | ✅ | 罫線あり、ヘッダー行太字 |
| 改ページ（`<!-- pagebreak -->`） | ✅ | HtmlBlock として検出 |
| 改ページ（`---pagebreak---`） | ✅ | ParagraphBlock として検出 |
| 日本語本文 | ✅ | EastAsia フォント名設定 |
| インライン太字（`**` / `__`） | ✅ | `RunProperties.Bold` |
| インライン斜体（`*` / `_`） | ✅ | `RunProperties.Italic` |
| インラインコード（`` ` ``） | ✅ | Courier New フォント |
| 太字＋斜体（`***`） | ✅ | 両方の RunProperties を適用 |
| `<br>` / ソフトリターン | ✅ | `LineBreakInline` / `HtmlInline` → `Break` |
| 水平線（`---`） | ✅ | `ParagraphBorders.BottomBorder` |
| ヘッダー文字列 | ✅ | `HeaderPart` + `SectionProperties` |
| フッター（ページ番号） | ✅ | `FieldChar` + `FieldCode(" PAGE ")` |
| 見出し番号付与 | ✅ | `numberHeadings=true` 時に文字列プレフィックス |
| 画像 | ❌ | 未対応 |
| 脚注・引用ブロック | ❌ | 未対応 |

---

## 4. テスト結果（v0.5.2 時点）

`tests/Md2Doc.Tests/OpenXmlConverterTests.cs` に 26 テストを実装。
**すべて Microsoft Word 不要で実行可能（CI で実行される）。**

### 要素別テスト（v0.5.1 から継続）

| テスト名 | 観点 |
|----------|------|
| `Headings_H1ToH3_AllPresent` | H1〜H3 すべて出力 |
| `Paragraph_JapaneseText_Preserved` | 日本語テキスト保全 |
| `BulletList_AllItemsPresent` | 4 箇条書きの欠落なし |
| `OrderedList_AllItemsPresent` | 順序付きリスト |
| `OrderedList_MultipleListsNumberIndependently` | 複数リストの独立番号 |
| `Table_HeaderAndDataCellsPresent` | テーブル全セル |
| `PageBreak_HtmlComment_BothSidesPreserved` | `<!-- pagebreak -->` |
| `PageBreak_DashSyntax_BothSidesPreserved` | `---pagebreak---` |
| `Headings_NumberingEnabled_PrefixAdded` | 見出し番号付与 |

### v0.5.2 新機能テスト

| テスト名 | 観点 |
|----------|------|
| `Header_TextPresent` | ヘッダー文字列が header パートに出力 |
| `Footer_PageNumber_Present` | フッターに PAGE フィールドが出力 |
| `HorizontalRule_Present` | 水平線（段落下罫線）が出力 |
| `InlineFormatting_Bold_Present` | 太字 `<w:b/>` が Run に設定 |
| `InlineFormatting_Italic_Present` | 斜体 `<w:i/>` が Run に設定 |
| `InlineFormatting_BoldAndItalicCombined` | 太字＋斜体の同時適用 |
| `InlineFormatting_InlineCode_UsesCodeFont` | インラインコードが Courier New |
| `SoftReturn_HardBreak_Preserved` | 行末2スペース → `w:br` |
| `SoftReturn_BrTag_Preserved` | `<br>` タグ → `w:br` |

### 回帰テスト（v0.5.1 から継続）

| テスト名 | 観点 |
|----------|------|
| `Regression_AllTextPresentInOrder` | 回帰: 全テキスト順序 |
| `Regression_NoBulletAlternatingLoss` | 回帰: 交互消失なし |
| `Regression_OrderedAndUnorderedMixed` | 回帰: 混在リスト |
| `Regression_BulletsAroundTable` | 回帰: テーブル前後 |

**Word COM 版テスト（`DocxConversionTests`）と同一 Markdown・同一観点** で
`Regression_*` テストを設計しており、エンジン間の品質比較基準が揃っている。

---

## 5. POC → 本実装候補で確認された利点

### CI 完全統合

Word COM 方式では `DocxConversionTests` が Word なし環境でスキップ（passed 偽装）されていたが、
Open XML 方式のテストは Word なし環境でも実際に実行される。

### 挙動の確定性

Word COM 方式では v0.3.0 にて以下の問題が発生した：
- 段落再計算タイミングによる本文消失
- `ApplyBulletDefault()` のトグル動作による箇条書き消失
- 隣接段落へのリスト書式自動伝播

Open XML 方式はこれらの問題が構造的に発生しない。

### インライン書式（v0.5.2 で対応）

v0.5.1 では除去のみだったインライン書式を、v0.5.2 で `RunProperties` を通じて完全対応した。
`EmphasisInline.DelimiterCount` で太字（≥2）と斜体（=1）を区別し、
`CodeInline` は等幅フォント（Courier New）を適用する。

### ヘッダー・フッター（v0.5.2 で対応）

`HeaderPart` / `FooterPart` + `SectionProperties` の三点セットで実装。
ページ番号は OOXML フィールド（`FieldChar` + `FieldCode`）方式で実装した。

---

## 6. 残存制約・懸念点

### インライン書式（リスト項目）

本文段落・見出し・表セルのインライン書式は対応済み。
リスト項目のインライン書式（例: `- **太字**項目`）は、lazy continuation 分割処理の
制約により現状はプレーンテキスト出力となる。将来対応可能。

### 画像・脚注・引用ブロック

POC / 本実装候補のスコープ外。必要に応じて追加実装する。

### 配布サイズの増加

`Markdig` + `DocumentFormat.OpenXml` の追加で実行ファイルのサイズが約 6〜8 MB 増加する。

### スタイル定義の最小実装

見出しスタイルは Simple な定義のみ。Word の既定テンプレートに比べて
色・間隔等が異なる場合がある。完全互換には Word の OOXML テンプレートに近い
スタイル定義が必要（v0.6.0 以降で対応）。

---

## 7. v0.6.0 以降の課題

| 作業 | 優先度 | コスト目安 |
|------|--------|-----------|
| 正式エンジン採用判断・WordInteropConverter deprecate | 高 | 設計判断 |
| Word COM 方式廃止タイミングの決定 | 高 | 設計判断 |
| リスト項目のインライン書式対応 | 中 | 〜80 行 |
| スタイル定義の完全化（色・間隔） | 中 | 〜100 行 |
| 画像対応（`ImageInline` → `ImagePart`） | 低 | 〜150 行 |

---

## 8. v0.5.3 実機確認結果

v0.5.3 で変換方式切替 UI を追加し、実機での比較確認を行った。

| 確認項目 | 結果 |
|---------|------|
| 本文消失（v0.3.0 の既知問題） | Open XML 方式では再現しなかった |
| 段落番号付与の不備 | Open XML 方式では再現しなかった |
| 変換速度（Word COM 方式） | 約 10 秒程度 |
| 変換速度（Open XML 方式） | 1 秒未満・体感 0.1 秒程度 |
| 移行可否の見込み | Open XML 方式へ移行可能と判断できる見込み |

変換速度の差異は、Word COM 方式が Word プロセス起動 + COM Marshal のオーバーヘッドを
伴うのに対し、Open XML 方式が純粋な XML 構築のみで完結するためである。

---

## 9. フォント設定の仕様整理（v0.5.4）

Open XML 方式における文書フォント設定の仕様を以下のとおり定める。

| 対象 | 仕様 |
|------|------|
| 本文フォント種別 | 文書フォント設定の指定値を適用 |
| 見出しフォント種別 | 文書フォント設定の指定値を適用（本文と同一） |
| 見出しフォントサイズ | H1: 16pt / H2: 14pt / H3: 12pt（スタイル定義に従い固定） |
| 見出しフォントサイズの個別指定 | 現時点では対象外 |

Word COM 方式では v0.3.0 にて見出しフォントの上書きを廃止し、Word の標準スタイル定義に
委ねる設計に統一した。Open XML 方式も同方針とする。
フォント設定が見出しにも反映されるため、UI のグループボックスラベルを
「本文フォント設定」から「文書フォント設定」に変更した（v0.5.4）。

---

## 10. 結論

v0.5.1 POC → v0.5.2 本実装候補化 → v0.5.3 UI 統合 → v0.5.4 実機確認を経て、
Open XML 方式は **現行 WordInteropConverter の代替として十分な品質に達した** と評価できる。

- 必須 Markdown 要素（見出し・段落・箇条書き・順序付きリスト・表・改ページ・日本語・
  インライン書式・ヘッダー・フッター・水平線・ソフトリターン）を網羅
- Word COM 方式で発生した挙動依存バグは構造的に排除
- 26 テストが CI で完全実行され、回帰検知の品質が大幅に向上
- 実機確認で既知問題の非再現・変換速度の大幅改善を確認

v0.6.0 で Open XML 方式を標準方式とする安定版候補に進める。
Word COM 方式は互換確認用として引き続き維持する。
