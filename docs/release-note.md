# Release Notes

## v0.6.0

title: Open XML 方式を標準方式とする安定版候補（予定）

v0.5.4 の実機確認結果を受け、Open XML 方式を標準方式として確定し安定版候補に進める。

### 想定変更内容（未実装）

- Open XML 方式の標準方式化（初期値維持・Word COM 方式は互換確認用として存続）
- v1.0.0 に向けた `docs/test_scenarios.md` 全項目確認・残存課題の解消
- 配布パッケージの整備・README 更新

### 対象外（このバージョンでは扱わない）

- Word COM 方式の削除（互換確認用として維持）
- 画像・脚注・引用ブロック対応

---

## v0.5.4

title: 実機確認結果の反映・Open XML方式への移行方針整理

### 変更内容

- `MainForm.cs`: フォント設定グループボックスラベルを「本文フォント設定」→「文書フォント設定」に変更
  - 文書フォント設定は本文・見出し両方のフォント種別に反映される仕様のため
- `docs/poc-openxml.md`:
  - v0.5.3 実機確認結果を記録（本文消失の非再現・変換速度比較・移行可否判断）
  - フォント設定の仕様を整理（フォント種別は本文・見出し共通、見出しサイズはスタイル定義固定）
  - 結論を v0.5.4 実機確認を踏まえた内容に更新
- `docs/engine-comparison.md`:
  - v0.5.4 実機確認結果を比較表で記録
  - 推奨を更新: v0.6.0 で Open XML 方式を標準方式とする安定版候補に進める方針を明記
- `docs/test_scenarios.md`:
  - v0.5.3 実機確認結果を記録
  - v1.0.0 配布向け確認項目（6-1〜6-4）を追加
- `docs/release-note.md`: v0.5.4 エントリを追加・v0.6.0 定義を更新

### フォント設定仕様の整理

| 対象 | フォント種別 | フォントサイズ |
|------|------------|-------------|
| 本文 | 文書フォント設定の指定値 | 文書フォント設定の指定値 |
| 見出し H1 | 文書フォント設定の指定値 | 16pt（スタイル定義固定） |
| 見出し H2 | 文書フォント設定の指定値 | 14pt（スタイル定義固定） |
| 見出し H3 | 文書フォント設定の指定値 | 12pt（スタイル定義固定） |

### 対象外（実装しない）

- Open XML 方式の正式採用・Word COM 方式の削除（v0.6.0 以降）
- 見出しフォントサイズの個別指定
- 画像・脚注・引用ブロック対応

### 既知の制約

- v0.5.3 の制約をすべて引き継ぐ

---

## v0.5.3

title: 変換方式選択 UI の追加・Open XML 方式の比較利用開始

### 変更内容

- オプション欄に「変換エンジン」選択行を追加
  - **Open XML 方式（Word不要・標準候補）** — 初期値
  - **Word COM 方式（互換確認用・Microsoft Word必要）**
- 選択したエンジンに応じて `OpenXmlConverter.ConvertToDocx()` / `WordInteropConverter.ConvertToDocx()` を切り替えて呼び出す
- 変換結果ラベルに使用エンジンを表示（例: `変換完了 [Open XML]: output.docx`）
- 選択したエンジンを `ConversionEngine` として設定ファイルに保存・復元
  - 既存設定ファイルに `ConversionEngine` がない場合は `"OpenXml"` を既定値とする
- フォームタイトルを「Markdown変換ツール（Word）」→「Markdown変換ツール」に変更
- `UserSettings.ConversionEngine` プロパティを追加

### 選択 UI の意図

- **Open XML 方式を初期値** としたのは、v1.0.0 での標準候補として検証を積むため
- Word COM 方式は削除せず **互換確認用** として維持する
- 両方式で同一入力・同一設定値（フォント・ヘッダー・フッター等）を共通利用する

### 対象外（実装しない）

- Open XML 方式の正式採用・Word COM 方式の廃止（v0.6.0 以降）
- 画像・脚注・引用ブロック対応
- 大規模な画面改修

### 既知の制約

- v0.5.2 の制約をすべて引き継ぐ
- Open XML 方式でのリスト項目内インライン書式はプレーンテキスト出力（v0.6.0 以降で対応予定）

---

## v0.5.2

title: Open XML 方式の本実装候補化

### 変更内容

- `src/md2doc/OpenXmlConverter.cs` を機能拡張（本実装候補クラスに格上げ）
  - **インライン書式**: 太字（`**` / `__`）・斜体（`*` / `_`）・インラインコード（`` ` ``）を `RunProperties` で装飾出力
  - **ヘッダー文字列**: `HeaderPart` + `SectionProperties.HeaderReference` で出力（配置指定対応）
  - **フッター（ページ番号）**: `FooterPart` + `FieldChar / FieldCode(" PAGE ")` で出力
  - **水平線（HR）**: `ThematicBreakBlock` → `ParagraphBorders.BottomBorder` で出力
  - **ソフトリターン改善**: `LineBreakInline`（硬改行）および `<br>` タグ → `w:br` で確実に出力
  - インラインレンダリングを `ExtractInlines`（文字列）から `AppendInlineRuns`（Run 要素直接生成）に刷新
- `tests/Md2Doc.Tests/OpenXmlConverterTests.cs` に 9 テストを追加（計 26 テスト）
  - `Header_TextPresent` / `Footer_PageNumber_Present` / `HorizontalRule_Present`
  - `InlineFormatting_Bold_Present` / `InlineFormatting_Italic_Present`
  - `InlineFormatting_BoldAndItalicCombined` / `InlineFormatting_InlineCode_UsesCodeFont`
  - `SoftReturn_HardBreak_Preserved` / `SoftReturn_BrTag_Preserved`
- `tests/Md2Doc.Tests/DocxInspector.cs` に書式・要素検査メソッドを追加
  - `HasRunWithTextAndProperty` / `HasRunWithFont` / `FooterHasPageNumber`
  - `DocumentHasHorizontalRule` / `DocumentHasSoftReturn` / `ExtractHeaderTexts`
- `docs/poc-openxml.md` を v0.5.2 時点の内容に全面更新

### 対応要素（v0.5.2 完了後）

| 要素 | 対応 |
|------|------|
| 見出し H1〜H3 / 通常段落 / 箇条書き / 順序付きリスト / 表 / 改ページ / 日本語 | ✅ v0.5.1 |
| インライン太字・斜体・インラインコード | ✅ v0.5.2 |
| ヘッダー文字列 / フッターページ番号 | ✅ v0.5.2 |
| 水平線 / ソフトリターン | ✅ v0.5.2 |
| 画像・脚注・引用ブロック | ❌ 未対応 |

### 対象外（実装しない）

- UI へのエンジン切替機能追加（v0.6.0 以降）
- 現行 Word COM エンジンの削除
- Open XML 方式の標準エンジン化

### 既知の制約

- v0.5.1 の制約のうち、インライン書式・ヘッダー・フッター・水平線・ソフトリターンは本バージョンで解消
- リスト項目内のインライン書式は lazy continuation 処理の制約によりプレーンテキスト出力のまま
- スタイル定義（見出し色・間隔）は最小実装のまま（v0.6.0 以降で対応予定）

---

## v0.5.1

title: Markdig + Open XML SDK 方式の POC 実装

### 変更内容

- `src/md2doc/OpenXmlConverter.cs` を新規追加（POC 変換クラス、Word COM 不要）
  - Markdig（AST パース）+ DocumentFormat.OpenXml（OOXML 直接構築）
  - WordInteropConverter と同一シグネチャで共存
  - 対応要素: 見出し H1〜H3 / 通常段落 / 箇条書き / 順序付きリスト / 表 / 改ページ / 日本語本文
  - 順序付きリストに正式対応（Word COM 方式では未対応だった）
  - v0.3.0 の本文消失・箇条書き伝播問題が構造的に発生しない
- `tests/Md2Doc.Tests/OpenXmlConverterTests.cs` を新規追加（13 テスト、全 Word 不要）
  - CI で完全実行可能（Word なし環境でも passed 偽装なし）
  - DocxConversionTests と同一 Markdown・同一観点で回帰テストを設計
- `docs/poc-openxml.md` を新規追加（POC 実装記録・制約・移行リスト）
- `docs/engine-comparison.md` に POC 結果を追記
- NuGet 追加: `Markdig 0.37.0`（BSD-2-Clause）/ `DocumentFormat.OpenXml 3.2.0`（MIT）

### 対象外（実装しない）

- 現行 Word COM エンジンの削除
- UI へのエンジン切替機能追加
- インライン書式の装飾（POC では除去のみ）
- ヘッダー・フッター対応
- 正式採用判断（v0.5.2 / v0.6.0 以降）

### 既知の制約

- v0.4.0 の制約をすべて引き継ぐ
- OpenXmlConverter は POC クラスであり、本番運用は現行 WordInteropConverter を使用する
- 配布サイズが Markdig + DocumentFormat.OpenXml の追加で約 6〜8 MB 増加

## v0.5.0

title: 変換エンジン比較検証（ドキュメントのみ）

### 変更内容

- `docs/engine-comparison.md` を新規作成
  - 現行 Word COM 方式 / Markdig + Open XML SDK / Markdig + HtmlToOpenXml の 3 方式を比較
  - 機能・運用・ライセンス各観点で評価表を整備
  - v0.5.0 時点での推奨案として「Markdig + Open XML SDK」を本命候補に位置付け
  - 採用判断のための次ステップ（POC 実装・既存テスト基盤での比較検証等）を提案
- 本リリースはコード変更を含まない。現行エンジンは温存し、UI も変更しない

### 対象外（実装しない）

- 現行 Word COM エンジンの削除
- UI へのエンジン切替機能追加
- 方式 B / C の本実装
- 外部 EXE 同梱方式の検証
- 有償ライブラリ前提の実装

### 既知の制約

- v0.4.0 の制約をすべて引き継ぐ
- 本ドキュメントの推奨はあくまで比較検証段階のもので、正式採用は v0.6.0 以降に判断する

## v0.4.0

title: docx XML 検査による回帰テスト基盤の追加

### 変更内容

- テストプロジェクト `tests/Md2Doc.Tests/` を新規追加（xUnit、`net8.0-windows`）
- `MarkdownParser` を `WordInteropConverter` から分離し、Word COM 不要の単体テストを可能に
  - `ParseBlocks` / `Block` / `BlockKind` を `MarkdownParser.cs` に移動
- 単体テスト `ParseBlocksTests`: Word なしで実行可能（8 テスト）
- 統合テスト `DocxConversionTests`: `.docx` を ZIP 展開して `word/document.xml` を直接検査
  - Word がない環境では自動スキップ
  - 段落消失・箇条書き交互消失・テーブル前後の箇条書き消失を自動検出（5 テスト）
- `DocxInspector` ヘルパー: Word COM 不要、`System.IO.Compression` のみで動作

### 対象外（実装しない）

- 箇条書き安定モード（記号 + インデント方式）
- 順序付きリスト対応
- Parser / Model / Renderer の大規模分離
- UI 変更

### 既知の制約

- v0.3.0 の制約をすべて引き継ぐ
- 統合テストは Windows + Microsoft Word がインストールされた環境でのみ実行可能

## v0.3.0

title: 見出しスタイル・番号付与・改ページ対応

### 変更内容

- 見出し（`#` / `##` / `###`）を Word の標準見出しスタイル（見出し 1 / 見出し 2 / 見出し 3）として出力するよう変更
  - フォント上書きを廃止し、Word のスタイル定義の書式（色・サイズ等）をそのまま適用
  - 見出しフォント設定 UI を削除し、本文フォント設定に一本化
- 見出しに番号を付与できるオプションを追加（初期値: オフ）
  - 番号体系: `1.` / `1.1` / `1.1.1` 形式でテキストプレフィックスとして付与
- Markdown 内の改ページ記法 `<!-- pagebreak -->` / `---pagebreak---` を Word 改ページに変換

### 既知の制約

- v0.2.0 の制約をすべて引き継ぐ
- 見出し番号の書式（区切り文字・桁数）は固定
- 箇条書き記号の種類は Word テンプレートの `wdStyleListBullet` 定義に依存
- 順序付きリスト（`1.` / `2.`）は通常段落として出力される（箇条書き非対応）
- 見出しフォント・色は Word テンプレート定義に依存（フォント上書き機能は廃止）

### 実装備考

- Word COM の段落再計算タイミング問題（見出し直後の本文消失）を解消するため、
  Markdown 全文を `doc.Content.Text` に一括書き込みする 4 フェーズ方式を採用
- 詳細な調査経緯・COM 挙動のメモは `docs/dev-notes-word-interop.md` を参照

## v0.2.0

title: UI/UX 改善（設定保存・出力先提案・変換後アクション・エクスプローラアイコン）

### 変更内容

- エクスプローラ表示用アイコンを追加（`app.ico` をビルドに埋め込み）
- 変換完了後に「Wordファイルを開く」「保存先フォルダを開く」ボタンを有効化
- 出力先パスを自動提案
  - テキスト入力: 最後に使った出力フォルダ（初回はデスクトップ）+ `output.docx`
  - ファイル入力: 入力ファイルと同じフォルダ・同じ名前で `.docx` 拡張子
- 以下の設定を `%LOCALAPPDATA%\md2doc\settings.json` に保存・次回起動時に復元
  - 見出し・本文フォントとサイズ
  - ヘッダー・フッター設定（種類・カスタムテキスト・配置）
  - 最後に使った出力フォルダ
  - ウィンドウサイズ

### 既知の制約

- v0.1.5 の制約をすべて引き継ぐ

## v0.1.5

title: 水平線対応・アイコン追加（検証版）

### 変更内容

- `---` / `***` / `___`（3文字以上）を Word の段落罫線（水平線）として出力するよう対応
- アプリケーションアイコンを追加（青背景に白文字「M」、32×32px・実行時生成）

### 既知の制約

- v0.1.4 の制約をすべて引き継ぐ

## v0.1.4

title: ドラッグ&ドロップ・テーブル対応（検証版）

### 変更内容

- テキストファイル（.md / .txt）のドラッグ&ドロップに対応（ファイル内容をテキスト入力エリアに読み込み、テキスト入力モードに自動切替）
- Markdown テーブル記法（`| col | col |` 形式）を Word テーブルとして出力するよう対応
  - 区切り行（`|---|---|`）がある場合は先頭行をヘッダーとして太字にする
  - 罫線あり・セルごとにフォント適用

### 既知の制約

- v0.1.3 の制約をすべて引き継ぐ

## v0.1.3

title: <br>対応・クリアボタン・進捗表示・レイアウト修正（検証版）

### 変更内容

- `<br>` / `<br/>` タグを Word のソフトリターン（段落内改行）として出力するよう対応
- Markdown テキストエリアにクリアボタンを追加（テキスト入力モード時のみ有効）
- 変換中の進捗率を「変換中... X%」形式で表示（Word 起動 5% → 行処理 10〜80% → 保存 95%）
- フォント設定の「本文」ラベルが別行に落ちるレイアウト崩れを根本修正

### 既知の制約

- v0.1.2 の制約をすべて引き継ぐ

## v0.1.2

title: フォント履歴・配置選択・作者名の追加（検証版）

### 変更内容

- フォント選択に「最近使用した直近5件」の履歴から選択できる機能を追加（`%LOCALAPPDATA%\md2doc\font_history.json` に保存）
- ヘッダーとフッターの文字配置を「左寄り」「中央」「右寄り」から選択可能に
- Word文書の作者名を「md2doc」として設定するよう変更

### 既知の制約

- v0.1.0 の制約をすべて引き継ぐ

## v0.1.0

title: 初回リリース（試作版）

### 変更内容

- WinForms + Office Interop による Markdown → Word (.docx) 変換
- テキスト入力 / ファイル入力の選択
- フォント名・フォントサイズの指定
- 上書き確認ダイアログ
- COM オブジェクトの適切な解放処理

### 既知の制約

- 対応 Markdown 要素: 見出し（H1〜H3）、箇条書き（`-` / `*`）、段落
- インライン書式（`**太字**`・`*斜体*`・`` `コード` ``）はマーカー除去のみ（文字装飾なし）
- 順序付きリスト・表・画像・脚注は非対応
- 実行には .NET 8 Desktop Runtime と Microsoft Word が必要
