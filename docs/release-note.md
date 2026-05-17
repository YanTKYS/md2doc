# Release Notes

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
