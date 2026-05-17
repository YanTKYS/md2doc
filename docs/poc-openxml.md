# POC 記録: Markdig + Open XML SDK 方式

v0.5.1 で実施した Markdig + Open XML SDK 方式（方式 B）の POC 実装記録。
`docs/engine-comparison.md` の推奨を受けて、実用可能かを検証する。

---

## 1. 実装概要

### 対象ファイル

| ファイル | 説明 |
|----------|------|
| `src/md2doc/OpenXmlConverter.cs` | POC 変換クラス（Word COM 不要） |
| `tests/Md2Doc.Tests/OpenXmlConverterTests.cs` | POC テスト（全 Word 不要） |

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
| COM 安定化 | 4 フェーズ + 2 パス補正 | 不要（直接 XML 構築） |

---

## 3. 対応要素

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
| インライン書式（太字・斜体） | △ | テキスト抽出のみ（装飾なし） |
| インラインコード（`` ` `` ） | △ | テキスト抽出のみ |
| `<br>` / ソフトリターン | △ | `LineBreakInline` → `Break` |
| 水平線（`---`） | ❌ | スキップ |
| 見出し番号付与 | ✅ | `numberHeadings=true` 時に文字列プレフィックス |
| ヘッダー・フッター | ❌ | POC 対象外 |

---

## 4. テスト結果

`tests/Md2Doc.Tests/OpenXmlConverterTests.cs` に 13 テストを実装。
**すべて Microsoft Word 不要で実行可能（CI で実行される）。**

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
| `Regression_AllTextPresentInOrder` | 回帰: 全テキスト順序 |
| `Regression_NoBulletAlternatingLoss` | 回帰: 交互消失なし |
| `Regression_OrderedAndUnorderedMixed` | 回帰: 混在リスト |
| `Regression_BulletsAroundTable` | 回帰: テーブル前後 |

**Word COM 版テスト（`DocxConversionTests`）と同一 Markdown・同一観点** で
`Regression_*` テストを設計しており、エンジン間の品質比較基準が揃っている。

---

## 5. POC で確認された利点

### CI 完全統合

Word COM 方式では `DocxConversionTests` が Word なし環境でスキップ（passed 偽装）されていたが、
Open XML 方式のテストは Word なし環境でも実際に実行される。

### 挙動の確定性

Word COM 方式では v0.3.0 にて以下の問題が発生した：
- 段落再計算タイミングによる本文消失
- `ApplyBulletDefault()` のトグル動作による箇条書き消失
- 隣接段落へのリスト書式自動伝播

Open XML 方式はこれらの問題が構造的に発生しない。
段落は `body.Append()` で確定的に追加され、リスト書式は
`NumberingInstance` として XML に直接記述されるため、
Word の自動補正・伝播挙動に依存しない。

### 順序付きリスト対応

Word COM 方式では未対応（通常段落として処理）だったが、
Open XML 方式では `NumberingFormat.Decimal` + `LevelText = "%1."` で
正しく実装できた。

---

## 6. POC で確認された制約・懸念点

### インライン書式の非対応（POC 段階の省略）

太字・斜体・インラインコードはテキスト抽出のみで装飾なし。
本実装への移行時には `RunProperties` に `Bold`・`Italic`・`Monospace` フォント等を設定する必要がある。

### ヘッダー・フッター未対応（POC 対象外）

`WordInteropConverter` では `SetHeader()` / `SetFooterPageNumbers()` が実装済みだが、
Open XML 方式では POC 段階でスキップした。
`HeaderPart` / `FooterPart` + `SectionProperties` で実装可能（追加コスト: 50〜80 行程度）。

### 配布サイズの増加

`Markdig` + `DocumentFormat.OpenXml` の追加で実行ファイルのサイズが約 6〜8 MB 増加する。
現行方式（Word 必須、DLL 追加なし）に比べてトレードオフがある。

### 水平線（HR）未実装

`ThematicBreakBlock` を現在スキップしている。
Word の罫線（`Borders[-3].LineStyle = 1`）相当は `ParagraphBorders` で実装可能。

### スタイル定義の最小実装

見出しスタイルは Simple な定義のみ。Word の既定テンプレートに比べて
色・間隔等が異なる場合がある。完全互換には Word の OOXML テンプレートに近い
スタイル定義が必要。

---

## 7. 本実装化に向けた追加実装リスト

POC から本実装（`WordInteropConverter` の代替）に昇格させるために必要な作業：

| 作業 | 優先度 | コスト目安 |
|------|--------|-----------|
| インライン書式（太字・斜体・コード）のテキスト装飾 | 高 | 〜50 行 |
| ヘッダー・フッター（ファイル名表示・ページ番号） | 高 | 〜80 行 |
| 水平線（HR） | 低 | 〜15 行 |
| スタイル定義の完全化（色・間隔） | 中 | 〜100 行 |
| `<br>` タグ → ソフトリターン（`Break` 型区別） | 中 | 〜20 行 |
| 単体テストの拡充（スタイル属性、XML 構造検査） | 中 | 〜100 行 |

---

## 8. 結論

Markdig + Open XML SDK 方式は **実用可能** であることを POC で確認した。

- 対象 Markdown 要素（見出し・段落・箇条書き・順序付きリスト・表・改ページ・日本語）は
  すべて正しく変換される
- Word COM 方式で発生した挙動依存バグは構造的に排除される
- テストが CI で完全実行され、回帰検知の品質が大幅に向上する

本実装への移行を v0.5.2 以降の課題として提案する。
