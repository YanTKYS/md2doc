# テストシナリオ

## 1. 前提

- 対象: WinForms版 Markdown変換ツール（Word）
- 最終更新: v0.4.0
- 補足: この環境ではWindows/Office実機確認ができないため、実機検証項目を分離して記録する。

## 2. テストケース一覧

| ID | 区分 | シナリオ | 期待結果 | 状態 |
| --- | --- | --- | --- | --- |
| T-01 | 正常系 | テキスト入力で変換実行 | `.docx`が保存される | 未確認（要実機） |
| T-02 | 正常系 | ファイル入力で変換実行 | `.docx`が保存される | 未確認（要実機） |
| T-03 | 正常系 | フォント名・サイズ指定 | 指定値で出力される | 未確認（要実機） |
| T-04 | 異常系 | 入力なしで実行 | エラー表示される | 実装確認済 |
| T-05 | 異常系 | 入力ファイル不存在 | エラー表示される | 実装確認済 |
| T-06 | 異常系 | 出力先フォルダ不存在 | エラー表示される | 実装確認済 |
| T-07 | 異常系 | 上書き確認で「いいえ」 | 変換中止になる | 実装確認済 |
| T-08 | 境界値 | フォントサイズ最小値8 | 正常保存される | 未確認（要実機） |
| T-09 | 境界値 | フォントサイズ最大値72 | 正常保存される | 未確認（要実機） |
| T-10 | 再実行系 | 同一入力で再実行 | 再生成される | 未確認（要実機） |
| T-11 | Interop | COM例外（Word起動失敗） | 失敗メッセージを表示しアプリが継続可能 | 未確認（要実機） |
| T-12 | Interop | 連続5回変換 | Wordプロセス残留なし | 未確認（要実機） |

## 3. 整形確認（Markdown品質）

- [x] 見出し、本文、表が読みやすく改行されている。
- [x] チェック項目は `- [ ]` / `- [x]` で統一した。
- [x] 1行圧縮表記がないことを確認した。

## 4. 自動テスト（v0.4.0 追加）

### テストプロジェクト

`tests/Md2Doc.Tests/` （xUnit、`net8.0-windows`、`dotnet test` で実行）

### 単体テスト: ParseBlocksTests

Word COM 不要。`MarkdownParser.ParseBlocks` の純粋関数テスト。

| テスト名 | 検証内容 |
|----------|----------|
| `ParseBlocks_RegressionMarkdown_AllBlocksPresent` | 必須 Markdown 全 8 ブロックが正しく解析される |
| `ParseBlocks_NoBulletsLost_ConsecutiveBullets` | 4 連続箇条書きが欠落なく解析される |
| `ParseBlocks_HeadingFollowedByBullet_BulletNotLost` | 見出し直後の箇条書きが消えない |
| `ParseBlocks_ParagraphAfterBullet_ParagraphNotLost` | 箇条書き後の通常段落が消えない |
| `ParseBlocks_NumberHeadings_PrefixAdded` | 見出し番号が正しく付与される |
| `ParseBlocks_PageBreak_DetectedCorrectly` | 改ページ記法が検出される |
| `ParseBlocks_Table_ParsedCorrectly` | テーブルが正しく解析される |
| `ParseBlocks_InlineMarkup_Stripped` | インライン記法が除去される |

### 統合テスト: DocxConversionTests

**Windows + Microsoft Word が必要。** 未インストール環境ではスキップ。
`.docx` を展開し `word/document.xml` を直接検査する。

| テスト名 | 検証内容 |
|----------|----------|
| `Regression_AllTextPresentInOrder` | 必須 Markdown の全テキストが順序通り存在する（段落消失・交互消失・通常段落消失を一括検出） |
| `Regression_NoBulletAlternatingLoss` | 4 連続箇条書きが奇数/偶数行だけに減らない |
| `Regression_ParagraphAfterBulletNotLost` | 箇条書き後の通常段落が消えない |
| `Regression_BulletsAroundTable` | テーブル前後の箇条書きが消えない |
| `Regression_NumberedHeadings_TextPresent` | 番号付き見出し有効時に全テキストが存在する |

### 実行方法

```
dotnet test tests/Md2Doc.Tests/Md2Doc.Tests.csproj
```

### 検査手法

`DocxInspector.ExtractParagraphTexts()` が `word/document.xml` の `<w:p>` ごとに
`<w:t>` を結合して段落テキスト一覧を返す。
`DocxInspector.VerifyOrder()` が期待テキストの存在と順序を検証する。

## 5. 実行確認結果記録

- この実行環境では `dotnet` が存在せず、`dotnet build` は未実施。
- この実行環境では Windows / Office がないため、Interop実行確認は未実施。
- 実機確認は、Windows + Office導入済み端末で実施する。
- ParseBlocksTests は Word 不要のため、Windows でなくても `dotnet test` で実行可能。
