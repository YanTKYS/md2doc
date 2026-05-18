# テストシナリオ

## 1. 前提

- 対象: WinForms版 Markdown変換ツール（md2doc）
- 最終更新: v0.5.6
- 補足: この環境ではWindows/Office実機確認ができないため、実機検証項目を分離して記録する。
- Open XML方式を標準方式として整理している。Word COM方式は互換確認用・退避用として維持する。

---

## 2. 自動テスト（Word 不要）

### テストプロジェクト

`tests/Md2Doc.Tests/` （xUnit、`net8.0-windows`、`dotnet test` で実行）

### 2-1. 単体テスト: ParseBlocksTests

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

### 2-2. Open XML 方式テスト: OpenXmlConverterTests

**Microsoft Word 不要（CI で完全実行）。** 26 テスト。

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
| `Header_TextPresent` | ヘッダー文字列が header パートに出力 |
| `Footer_PageNumber_Present` | フッターに PAGE フィールドが出力 |
| `HorizontalRule_Present` | 水平線（段落下罫線）が出力 |
| `InlineFormatting_Bold_Present` | 太字 `<w:b/>` が Run に設定 |
| `InlineFormatting_Italic_Present` | 斜体 `<w:i/>` が Run に設定 |
| `InlineFormatting_BoldAndItalicCombined` | 太字＋斜体の同時適用 |
| `InlineFormatting_InlineCode_UsesCodeFont` | インラインコードが Courier New |
| `SoftReturn_HardBreak_Preserved` | 行末2スペース → `w:br` |
| `SoftReturn_BrTag_Preserved` | `<br>` タグ → `w:br` |
| `Regression_AllTextPresentInOrder` | 回帰: 全テキスト順序 |
| `Regression_NoBulletAlternatingLoss` | 回帰: 交互消失なし |
| `Regression_OrderedAndUnorderedMixed` | 回帰: 混在リスト |
| `Regression_BulletsAroundTable` | 回帰: テーブル前後 |

### 2-3. Word COM 方式統合テスト: DocxConversionTests

**Windows + Microsoft Word が必要。**
Word が利用できない環境（CI 含む）では、テスト本体を実行せずそのまま `return` する。

> **注意**: xUnit はテスト本体を実行しなかった場合も **passed** と表示する。
> テスト出力（`--logger "console;verbosity=normal"`）に `[SKIPPED]` メッセージが出ている場合、
> Word なし環境のため本体は未実行。

| テスト名 | 検証内容 |
|----------|----------|
| `Regression_AllTextPresentInOrder` | 必須 Markdown の全テキストが順序通り存在する |
| `Regression_NoBulletAlternatingLoss` | 4 連続箇条書きが奇数/偶数行だけに減らない |
| `Regression_ParagraphAfterBulletNotLost` | 箇条書き後の通常段落が消えない |
| `Regression_BulletsAroundTable` | テーブル前後の箇条書きが消えない |
| `Regression_NumberedHeadings_TextPresent` | 番号付き見出し有効時に全テキストが存在する |

### 2-4. 実行方法

```
dotnet test tests/Md2Doc.Tests/Md2Doc.Tests.csproj
```

---

## 3. 実機確認シナリオ（v0.6.0 配布前検証）

Windows + .NET 8 Runtime 導入済みの端末で実施する。
Open XML 方式を中心に確認し、必要に応じて Word COM 方式と比較する。

### 3-1. アプリ起動確認

| ID | 確認内容 | 期待結果 | 状態 |
|----|---------|---------|------|
| S-01 | アプリが起動する | メイン画面が表示される | 未確認 |
| S-02 | 起動時の変換エンジン初期値 | 「Open XML方式」が選択されている | 未確認 |
| S-03 | 前回のフォント設定が復元される | 設定が保持されている | 未確認 |
| S-04 | 前回の変換エンジン選択が復元される | 設定が保持されている | 未確認 |
| S-05 | 設定ファイルがない状態での起動 | 既定値で正常起動する | 未確認 |

### 3-2. テキスト入力からの変換（Open XML 方式）

| ID | 確認内容 | 期待結果 | 状態 |
|----|---------|---------|------|
| T-01 | テキスト入力で変換実行 | `.docx` が保存される | 未確認 |
| T-02 | 変換完了メッセージに `[Open XML]` と表示される | 表示される | 未確認 |
| T-03 | 出力 `.docx` を Word で開ける | 正常に開ける | 未確認 |

### 3-3. ファイル入力からの変換（Open XML 方式）

| ID | 確認内容 | 期待結果 | 状態 |
|----|---------|---------|------|
| F-01 | ファイル入力で変換実行 | `.docx` が保存される | 未確認 |
| F-02 | 入力ファイルと同フォルダに出力先が自動提案される | 提案される | 未確認 |
| F-03 | `.md` / `.txt` どちらも読み込める | 正常に変換される | 未確認 |

### 3-4. Open XML 方式: Markdown 要素確認

| ID | 確認内容 | 確認方法 | 状態 |
|----|---------|---------|------|
| OX-01 | 見出し H1 / H2 / H3 が出力される | 自動テスト / 実機 | 自動テスト済 |
| OX-02 | 見出し番号付与が正しく動作する（H1→`1.`、H2→`1.1`） | 自動テスト / 実機 | 自動テスト済 |
| OX-03 | 通常段落が出力される | 自動テスト / 実機 | 自動テスト済 |
| OX-04 | 箇条書き（`-` / `*`）が出力される | 自動テスト / 実機 | 自動テスト済 |
| OX-05 | 順序付きリスト（`1.`）が出力される | 自動テスト / 実機 | 自動テスト済 |
| OX-06 | 複数の順序付きリストが独立番号で出力される | 自動テスト / 実機 | 自動テスト済 |
| OX-07 | 表が罫線付きで出力される | 自動テスト / 実機 | 自動テスト済 |
| OX-08 | 改ページ（`<!-- pagebreak -->`）が動作する | 自動テスト / 実機 | 自動テスト済 |
| OX-09 | ヘッダー文字列が出力される | 自動テスト / 実機 | 自動テスト済 |
| OX-10 | フッターのページ番号が出力される | 自動テスト / 実機 | 自動テスト済 |
| OX-11 | 太字 / 斜体 / インラインコードが装飾付きで出力される | 自動テスト / 実機 | 自動テスト済 |
| OX-12 | 水平線が出力される | 自動テスト / 実機 | 自動テスト済 |
| OX-13 | `<br>` / 行末2スペースがソフトリターンとして出力される | 自動テスト / 実機 | 自動テスト済 |
| OX-14 | 文書フォント設定（種別・サイズ）が本文に反映される | 実機 | 未確認 |
| OX-15 | 文書フォント設定（種別）が見出しに反映される | 実機 | 未確認 |
| OX-16 | 見出しフォントサイズが H1:16pt / H2:14pt / H3:12pt になる | 実機 | 未確認 |

### 3-5. Word での開封・編集確認

| ID | 確認内容 | 状態 |
|----|---------|------|
| WD-01 | 出力 `.docx` を Word で開ける | 未確認（要実機） |
| WD-02 | 見出しスタイルが Word の見出し書式として認識される | 未確認（要実機） |
| WD-03 | 箇条書き / 順序付きリストが Word のリストとして再編集できる | 未確認（要実機） |
| WD-04 | 表のセルを Word で編集できる | 未確認（要実機） |
| WD-05 | ヘッダー / フッターが Word のヘッダー/フッター領域に表示される | 未確認（要実機） |
| WD-06 | ページ番号フィールドが Word で正しく更新される | 未確認（要実機） |
| WD-07 | 太字 / 斜体 / インラインコードの書式が Word で確認できる | 未確認（要実機） |

### 3-6. Word COM 方式との比較確認

| ID | 確認内容 | 状態 |
|----|---------|------|
| CMP-01 | 同一 Markdown を両方式で変換し、出力内容が同等である | 未確認（要実機） |
| CMP-02 | 見出し・箇条書き・表・改ページの見た目が同等である | 未確認（要実機） |
| CMP-03 | ヘッダー・フッターが両方式で同等に表示される | 未確認（要実機） |
| CMP-04 | 変換速度の差異が許容範囲内と判断できる | v0.5.3 実機確認で確認済み ✅ |

### 3-7. 変換時間の確認

| ID | 確認内容 | 期待結果 | 状態 |
|----|---------|---------|------|
| SPD-01 | Open XML 方式の変換時間（通常文書） | 1 秒未満（体感 0.1 秒程度） | v0.5.3 実機確認済み ✅ |
| SPD-02 | Word COM 方式の変換時間（通常文書） | 約 10 秒程度 | v0.5.3 実機確認済み ✅ |

### 3-8. 設定・UI 確認

| ID | 確認内容 | 状態 |
|----|---------|------|
| CFG-01 | アプリを再起動後に変換エンジン選択が復元される | 未確認（要実機） |
| CFG-02 | アプリを再起動後にフォント設定が復元される | 未確認（要実機） |
| CFG-03 | 既存の設定ファイルがある場合に互換性が保たれる | 未確認（要実機） |
| CFG-04 | ヘッダー・フッター設定が再起動後も復元される | 未確認（要実機） |

### 3-9. 異常系・境界値

| ID | 区分 | シナリオ | 期待結果 | 状態 |
|----|------|---------|---------|------|
| E-01 | 異常系 | 入力なしで実行 | エラー表示される | 実装確認済 |
| E-02 | 異常系 | 入力ファイル不存在 | エラー表示される | 実装確認済 |
| E-03 | 異常系 | 出力先フォルダ不存在 | エラー表示される | 実装確認済 |
| E-04 | 異常系 | 上書き確認で「いいえ」 | 変換中止になる | 実装確認済 |
| E-05 | 境界値 | フォントサイズ最小値 8 | 正常保存される | 未確認（要実機） |
| E-06 | 境界値 | フォントサイズ最大値 72 | 正常保存される | 未確認（要実機） |
| E-07 | 再実行系 | 同一入力で再実行 | 上書き確認後に再生成される | 未確認（要実機） |
| E-08 | Interop | Word COM 方式: COM例外（Word起動失敗） | 失敗メッセージを表示しアプリが継続可能 | 未確認（要実機） |
| E-09 | Interop | Word COM 方式: 連続5回変換 | Wordプロセス残留なし | 未確認（要実機） |

### 3-10. 既知の制約の確認

| ID | 確認内容 | 期待結果 | 状態 |
|----|---------|---------|------|
| LIM-01 | 画像記法を含む Markdown を変換 | 画像は無視される（エラーにならない） | 未確認（要実機） |
| LIM-02 | H4 以上の見出しを含む Markdown を変換（番号付与あり） | H4 以上は番号なしで H3 スタイルとして出力される | 未確認（要実機） |
| LIM-03 | Word COM 方式で順序付きリストを変換 | 通常段落として出力される（番号なし） | 未確認（要実機） |

---

## 4. 整形確認（Markdown品質）

- [x] 見出し、本文、表が読みやすく改行されている。
- [x] チェック項目は `- [ ]` / `- [x]` で統一した。
- [x] 1行圧縮表記がないことを確認した。

---

## 5. 実行確認結果記録

- この実行環境では `dotnet` が存在せず、`dotnet build` は未実施。
- この実行環境では Windows / Office がないため、Interop 実行確認は未実施。
- 実機確認は、Windows + .NET 8 Runtime 導入済み端末で実施する。
- Open XML 方式の ParseBlocksTests / OpenXmlConverterTests は Word 不要のため、Windows でなくても `dotnet test` で実行可能。

### v0.5.3 実機確認結果（2026-05-18）

v0.5.3 で変換方式切替 UI を追加後に実機確認を行った。

| 確認項目 | 結果 |
|---------|------|
| Open XML 方式: 本文消失の再現 | 再現しなかった ✅ |
| Open XML 方式: 段落番号付与不備の再現 | 再現しなかった ✅ |
| Word COM 方式: 変換速度 | 約 10 秒程度 |
| Open XML 方式: 変換速度 | 1 秒未満・体感 0.1 秒程度 ✅ |
| Open XML 方式への移行可否 | 移行可能と判断できる見込み ✅ |

---

## 6. バージョン体系と今後の方針

| バージョン | 位置づけ |
|----------|---------|
| v0.5.x | Open XML 方式への移行準備シリーズ |
| v0.6.x | 配布前検証・受入テスト・運用整備フェーズ |
| v1.0.0 | 庁内配布可能版（目標） |

**v0.6.0 以降の実機確認では「3. 実機確認シナリオ」のチェックリストを使用する。**
Open XML 方式を標準方式として確認を進め、Word COM 方式は互換確認・退避用として参照扱いとする。
