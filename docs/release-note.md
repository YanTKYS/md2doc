# Release Notes

## v0.6.2

title: 設定サンプル表示機能の削除

v0.6.0 で追加した設定サンプル表示は、変換後の Word 文書の見え方を再現するものではなく、
ヘッダー・フッター・見出し番号の部品表示に留まるものであった。
実機確認の結果、UX 上の価値が低いと判断し、機能ごと削除する。
変換後は「Wordファイルを開く」ボタンから Word で直接確認する導線を優先する。

### 変更内容

- **`src/md2doc/MainForm.cs`**
  - 設定サンプル表示 GroupBox を削除（`BuildSampleGroupBox()`・`UpdateSample()` メソッド削除）
  - `_sampleUpdateButton` / `_sampleTextBox` / `_sampleNoteLabel` フィールドを削除
  - `_sampleUpdateButton.Click` イベントハンドラを削除
  - 設定セクションを 5 行 → 4 行に縮小（sampleGroupBox 行の除去）
- **`src/Md2Doc.Core/SettingsSampleBuilder.cs` を削除**
- **`tests/Md2Doc.Tests/SettingsSampleBuilderTests.cs` を削除**

### 維持される動作

- Open XML 方式・Word COM 方式の変換処理
- フォント・ヘッダー・フッター・見出し番号設定
- 変換後の「Wordファイルを開く」「保存先フォルダを開く」ボタン
- v0.6.1 のレイアウト修正（2 段構成・`MinimumSize`）

### 既知の制約

- v0.6.1 の制約をすべて引き継ぐ

---

## v0.6.1

title: v0.6.0 レイアウト崩れの修正

v0.6.0 で追加した設定サンプル表示に伴う画面レイアウト崩れを修正する。
変換ロジック・設定サンプル生成ロジックに変更はない。

### 問題

- Markdown 本文入力欄が極端に潰れて表示されない
- 設定サンプル表示欄（GroupBox）が画面の大半を占有し、入力欄を圧迫する
- 画面サイズを変更した際に UI が崩れる
- フォームに最小サイズが設定されていない

### 原因

v0.6.0 では、単一の 13 行 `TableLayoutPanel` に入力欄（`Percent,100` 行）と
設定 GroupBox 群（すべて `AutoSize` 行）を混在させていた。
`BuildSampleGroupBox()` が内部で `TableLayoutPanel` を使い
絶対位置指定（`Location`）を行ったため、GroupBox の `AutoSize` 算出が
コントロールの実サイズを反映せず、`Percent,100` 行に割り当てられるべき
スペースが正しく計算されなかった。

### 修正内容

- **レイアウトを 2 段構成に変更**（`MainForm.cs`）
  - 上部：入力セクション（`Percent,100`）— Markdown テキスト・ファイル入出力
  - 下部：設定セクション（`AutoSize`）— フォント・オプション・サンプル・ボタン
  - 設定セクション側の AutoSize 行が入力欄の伸縮スペースを奪わない
- **`BuildSampleGroupBox()` を修正**
  - 内部コンテナを `TableLayoutPanel` + 絶対位置指定 から
    `FlowLayoutPanel (FlowDirection.TopDown)` に変更し、
    フォント GroupBox と同一パターンに統一
  - `_sampleTextBox.Width = 840` を明示して GroupBox の AutoSize 算出を安定化
- **`MinimumSize = new Size(900, 750)` を設定**
  - 画面を縮めすぎた場合に主要 UI が重なるのを防止
- **`_markdownTextBox.MinimumSize = new Size(100, 80)` を設定**
  - 最小高さを保証し、ウィンドウが最小サイズ近辺でも入力欄が視認可能
- **`_sampleTextBox.Height = 100`（120 → 100）**
  - 設定セクションの占有高さを抑制

### 維持される動作

- v0.6.0 のすべての機能（設定サンプル表示・変換処理）
- 変換エンジン選択・設定保存復元
- フォント設定・ヘッダー・フッター設定

### 既知の制約

- v0.6.0 の制約をすべて引き継ぐ
- 設定サンプル表示欄の横幅は 840px 固定（ウィンドウを極端に広くしても伸長しない）

---

## v0.5.7

title: Open XML 方式の Core 切り出し・テスト構成整理

v0.6.0 で Open XML 方式を標準方式の安定版候補へ進める前段として、変換処理を
WinForms 本体から独立したライブラリプロジェクトへ切り出し、責務を分離する。
新機能追加・UI 変更・Word COM 方式廃止は行わない。

### 変更内容

- **`src/Md2Doc.Core/` プロジェクトを追加**
  - `TargetFramework=net8.0`、WinForms 非依存・Microsoft Word 非依存
  - `Markdig` / `DocumentFormat.OpenXml` を参照
- **`OpenXmlConverter` を Core プロジェクトへ移動**
  - 名前空間: `Md2Doc` → `Md2Doc.Core`
  - 可視性: `internal static` → `public static`（ライブラリ API として明示）
  - 実装ロジックは変更なし
- **WinForms 本体（`src/md2doc/`）から Open XML 関連の依存を除去**
  - `md2doc.csproj` から `Markdig` / `DocumentFormat.OpenXml` の直接参照を削除
    （`Md2Doc.Core` 経由で利用するため）
  - `MainForm.cs` に `using Md2Doc.Core;` を追加
- **テストプロジェクトを Core 参照に整理**
  - `Md2Doc.Tests.csproj` が `Md2Doc.Core` を直接参照
  - Open XML 方式の 26 テスト（`OpenXmlConverterTests`）は Core を対象に検証
  - Word COM 方式テスト（`DocxConversionTests`）と `ParseBlocksTests`、
    docx 検査ヘルパー `DocxInspector` は WinForms 本体側を引き続き参照
- **ソリューションファイル `md2doc.sln` を追加**
  - 3 プロジェクト（`Md2Doc.Core` / `md2doc` / `Md2Doc.Tests`）を一括ビルド可能に
- **`build.yml` の更新**
  - `dotnet build md2doc.sln` でソリューション全体をビルド
  - テスト実行は `--no-build` で重複ビルドを回避
- **`release.yml` は変更なし**
  - `dotnet publish src/md2doc/md2doc.csproj` から `Md2Doc.Core` も
    ProjectReference 経由で自動含有される（`PublishSingleFile=true` により
    実行可能ファイル一本に統合される）

### 責務分離後の構成

| プロジェクト | 役割 | 主な依存 |
|------------|------|---------|
| `Md2Doc.Core` | Open XML 方式の変換処理（ライブラリ） | Markdig / DocumentFormat.OpenXml |
| `md2doc` (WinForms) | 画面表示・入力取得・出力先指定・設定保存復元・変換方式選択・呼び出し | Md2Doc.Core / Office Interop (dynamic) |
| `Md2Doc.Tests` | 単体・統合テスト | Md2Doc.Core / md2doc (Word COM 用) |

### 維持される動作

- Open XML 方式での変換（既定）
- Word COM 方式での変換（互換確認用）
- 変換方式選択 UI（起動時に Open XML 方式が初期選択）
- 設定保存・復元（フォント・ヘッダー・フッター・変換エンジン等）
- ヘッダー・フッター・ページ番号・フォント・見出し番号が Open XML 方式に反映される
- Open XML 方式の既存テスト 26 件・Word COM 方式の統合テスト 5 件・ParseBlocksTests 8 件

### Core 切り出しによる効果

- **テスト容易性**: Open XML 方式のテストが WinForms / Microsoft Word いずれにも
  依存せず、`Md2Doc.Core` 単体で検証可能
- **保守性**: 変換処理と画面処理の責務が明確に分離され、互いに影響しにくくなる
- **将来拡張性**: Open XML 方式のロジックが独立したアセンブリとなり、将来的に
  他フロントエンドから参照可能（v1.1.0 以降で必要性が確認された場合）

### 対象外（このバージョンでは扱わない）

- Word COM 方式の廃止
- Open XML 方式のみへの一本化
- CLI 作成・PowerShell 簡易版作成（`docs/backlog.md` に整理済み）
- 複数ファイルドラッグ＆ドロップ変換
- 画像・脚注・引用ブロック対応
- 大規模な UI 改修
- リリース作成

### 既知の制約

- v0.5.6 の制約をすべて引き継ぐ

---

## v0.5.5

title: 潜在的バグ修正・コード整理（メンテナンスリリース）

ソースコードベースのコードレビューで検出した潜在的バグの修正と、内部実装の整理を行う。
外部仕様・UI 文言・設定ファイル互換性に影響する変更はない。

### バグ修正

- `OpenXmlConverter.RenderHeading`: H4 以上の見出しが見出し番号カウンタを破壊する問題を修正
  - 修正前: H4 以上は H3 として描画され、H3 カウンタを誤って加算（例: H1 直後の H4 で `1.0.1` のような不正な番号が出力されていた）
  - 修正後: H4 以上はスタイルのみ H3 として描画し、番号付与とカウンタ更新の両方をスキップする
- `OpenXmlConverter.RenderHorizontalRule`: 水平線段落に空の Run を追加
  - Run なし段落は Word では正しく描画されるが、一部のリーダー（LibreOffice 等）で罫線が描画されない問題への対策
- `AppLog.Write`: ファイル書き込みを `lock` で直列化
  - `ConvertAsync` の `Task.Run`（バックグラウンドスレッド）と UI スレッドからの並行ログ書き込みで `IOException` が発生する可能性を排除
- `UserSettings.Load`: 不正な設定値の正規化処理を追加
  - 手動編集や過去バージョンとの非互換に備え、`HeaderMode` / `HeaderAlignment` / `FooterAlignment` / `ConversionEngine` / `BodyFontSize` の範囲外値を既定値にクランプする

### リファクタリング

- `WordInteropConverter`: Word VBA 定数（`wdStyleNormal=-1` 等）をマジックナンバーから名前付き `const int` に置換
  - `WdStyleNormal` / `WdStyleListBullet` / `WdBorderBottom` / `WdLineStyleSingle` / `WdListNoNumbering` / `WdCollapseEnd` / `WdFieldPage` / `WdFieldNumPages` / `WdFormatXMLDocument`
  - 見出しスタイルは `WdStyleHeading(level)` ヘルパー（`-(level+1)`）で算出

### 対象外（このバージョンでは扱わない）

- Word COM 方式と Open XML 方式の機能差異解消（順序付きリスト・`_em_` 構文・インラインコード等の差異は v0.6.0 以降で再評価）
- 9 引数の `ConvertToDocx` シグネチャを `ConversionOptions` レコードに集約する大規模リファクタリング（テスト 9 箇所への波及があり別途実施）

### 既知の制約

- v0.5.4 の制約をすべて引き継ぐ

---

## v0.5.6

title: Open XML 方式を標準前提としたドキュメント整理

v0.5.1〜v0.5.5 の実装・実機確認・品質整備を経て、Open XML 方式を今後の標準方式として
扱う前提でドキュメント全体の表現を整理する。コードの大きな変更は行わない。

### 変更内容

- `README.md`: 全面更新
  - Open XML 方式を標準方式として明記（Word 不要・高速変換）
  - Word COM 方式を互換確認用・退避用として整理
  - 対応 Markdown 要素・既知の制約・フォント仕様の一覧化
  - バージョン体系（v0.5.x / v0.6.x / v1.0.0）の明記
- `docs/engine-comparison.md`:
  - Section 8 を v0.5.5 / v0.5.6 状態に更新
  - Open XML 方式の標準採用決定を明記
  - 方式 A（Word COM）は互換確認用として維持する旨を明記
- `docs/test_scenarios.md`: 全面更新
  - Open XML 方式を中心とした v0.6.0 配布前検証シナリオを整備
  - 自動テスト（ParseBlocks / OpenXmlConverter / DocxConversion）一覧を整理
  - 実機確認シナリオ（起動・変換・Markdown 要素・Word 開封・比較・設定・異常系・制約確認）を体系化
  - バージョン体系と今後の方針を記載
- `manuals/user_manual.md`: Open XML 方式を前提とした手順に更新
  - 変換エンジン説明（Open XML 標準 / Word COM 互換確認用）を追加
  - フォント設定が本文・見出し両方に反映される旨を明記
- `manuals/operator_manual.md`: Open XML 方式前提の一次対応手順に更新
- `docs/release-note.md`: v0.5.6 エントリ追加・v0.6.0 定義を配布前検証フェーズとして更新

### フォント仕様の整理（確定）

| 対象 | フォント種別 | フォントサイズ |
|------|------------|-------------|
| 本文 | 文書フォント設定の指定値 | 文書フォント設定の指定値 |
| 見出し H1 | 文書フォント設定の指定値（本文と同一） | 16pt（スタイル定義固定） |
| 見出し H2 | 文書フォント設定の指定値（本文と同一） | 14pt（スタイル定義固定） |
| 見出し H3 | 文書フォント設定の指定値（本文と同一） | 12pt（スタイル定義固定） |

> 見出しフォントサイズの個別指定は現時点では対象外。

### 対象外（このバージョンでは扱わない）

- Word COM 方式の削除（互換確認用として維持）
- 画像・脚注・引用ブロック対応
- 見出しフォントサイズの個別指定
- Open XML 方式と Word COM 方式の機能差異解消
- 大規模な UI 改修・変換エンジンの再設計

### 既知の制約

- v0.5.5 の制約をすべて引き継ぐ

---

## v0.6.0

title: 設定サンプル表示機能の追加

オプション設定（ヘッダー・見出し番号・ページ番号）の確認を手軽にできるよう、
変換前に現在の設定をテキスト形式でプレビューできる「設定サンプル表示」機能を追加する。
変換処理そのものには変更を加えない。

### 変更内容

- **`src/Md2Doc.Core/SettingsSampleBuilder.cs` を新規追加**
  - 設定値（ヘッダーモード・テキスト・配置、見出し番号、ページ番号）を受け取り、
    人間が読みやすいサンプルテキストを返す `public static class`
  - Word 出力との完全一致は保証しない（設定確認用途に限定）
  - WinForms / Microsoft Word いずれにも非依存（`net8.0` Core ライブラリに配置）

- **`src/md2doc/MainForm.cs` を更新**
  - GroupBox「設定サンプル表示」を追加（オプション GroupBox の直下）
    - 「サンプル更新」ボタン: 押下時に現在のオプション設定を読み取り、サンプルテキストを更新
    - 読み取り専用 TextBox（マルチライン・スクロールバーあり）: サンプルテキストを表示
    - 免責ラベル: 「※この表示は設定確認用のサンプルです。実際のWord出力とは完全には一致しません。」
  - 変換処理・変換エンジン・フォント設定には変更なし

- **`tests/Md2Doc.Tests/SettingsSampleBuilderTests.cs` を新規追加**
  - ヘッダーなし / 自由記入（左・中央・右）/ ファイル名 / 未入力・未選択のケース
  - 見出し番号オン・オフ
  - ページ番号オン・オフ
  - セクション見出し（`【ヘッダー】`・`【見出し番号】`・`【ページ番号】`）の常時出力を確認

### サンプル表示の出力例

```
【ヘッダー】
左：業務手順書
中央：（なし）
右：（なし）

【見出し番号】
1. 大見出し
1.1 中見出し
1.1.1 小見出し

【ページ番号】
1 / 5
```

### 対象外（このバージョンでは扱わない）

- Word COM 方式の削除（互換確認用として維持）
- 画像・脚注・引用ブロック対応
- 見出しフォントサイズの個別指定
- フォントや余白を反映したリッチなプレビュー（WebView2 等の重い依存は追加しない）

### 既知の制約

- v0.5.7 の制約をすべて引き継ぐ
- サンプル表示は設定確認専用であり、実際の Word 出力と完全一致しない

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
