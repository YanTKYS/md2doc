# Markdown to Word 変換ツール

## 1. 概要

Markdown形式のテキスト、またはMarkdownファイル（`.md` / `.txt`）を  
Microsoft Word形式（`.docx`）へ変換する WinForms ツールです。

## 2. 目的

- 一般職員が、Markdown文書を業務配布しやすいWord文書へ変換できるようにする。
- 出力文書のフォント名・フォントサイズを指定し、組織内ルールに合わせる。

## 3. 対象利用者

- 一般職員
- 運用担当部署

## 4. 利用環境

- Windows端末
- Microsoft Office Standard導入済み（Word利用可能）
- .NET 8 Runtime / SDK（開発・ビルド時）
- 閉域ネットワーク（外部通信なし）

## 5. できること

- テキスト入力またはファイル入力でMarkdownを取り込む。
- `.docx`出力先を指定する。
- フォント名、フォントサイズを指定して出力する。
- 既存ファイル上書き時に確認ダイアログを表示する。

## 6. できないこと

- `.docx`以外（PDF/HTML等）への出力
- 画像URLの取得や外部リンク先の自動参照
- クラウド保存や外部API連携

## 7. 使い方（基本手順）

1. アプリを起動する。
2. 入力方式（テキスト入力 / ファイル入力）を選択する。
3. Markdown本文、または入力ファイルを指定する。
4. 出力ファイル（`.docx`）を指定する。
5. フォント名・フォントサイズを指定する。
6. `変換実行` を押す。
7. 結果表示を確認する。

## 8. 入力データ / 出力データ

- 入力
  - Markdown文字列
  - Markdownファイル（`.md` / `.txt`）
- 出力
  - Wordファイル（`.docx`）

## 9. 個人情報の取扱い

- 入力内容を外部送信しない。
- 個人情報を永続ログとして保存しない。
- 出力ファイルの保管・配布は利用部署の運用ルールに従う。

## 10. 注意事項（WinForms / Office Interop）

- Wordが起動できない場合、Office導入状態とライセンス状態を確認する。
- Office更新中や他プロセス競合時は変換失敗する場合がある。
- 変換中はWord COMオブジェクトを作成するため、異常終了時はWordプロセス残留を確認する。

## 11. 動作確認方法

```bash
dotnet build src/md2doc/md2doc.csproj
```

- `dotnet` がない環境では、Windows開発端末で確認する。

## 12. バージョン / 更新履歴 / 既知の制約 / 保守範囲

- バージョン区分: `v0.1.0`（試作版）
- 更新履歴:
  - 2026-05-05: 初期実装（WinForms + Office Interop）
- 既知の制約:
  - Markdown対応は見出し・箇条書き・段落・簡易インライン整形のみ。
  - 高度なMarkdown（表、画像、脚注等）は未対応。
- 保守範囲:
  - DX担当: ツール改修、配布、障害切り分け
  - 利用部署: 日常運用、一次確認
  - 所管部署判断: 出力文書の業務最終判断

## 13. 標準成果物と配置

本リポジトリは `reference/guide_context.md` の標準成果物に沿って、以下を配置しています。

- `README.md`
- `development_report.md`
- `docs/tool_design.md`
- `docs/release_checklist.md`
- `docs/test_scenarios.md`
- `docs/operation_handover.md`
- `manuals/admin_manual.md`
- `manuals/operator_manual.md`
- `manuals/user_manual.md`
- `src/md2doc/`（C# WinForms 実装）
