# 開発報告書

## 1. 今回の目的

既存の WinForms / Office Interop 実装を維持したまま、
同梱ガイド準拠（`reference/guide_context.md`、`reference/11_non_web_tool_patterns.md`、`reference/12_office_interop_checklist.md`）で
不足していた文書補完と確認記録を明確化する。

## 2. 参照した同梱ガイド

- `reference/guide_context.md`
- `reference/11_non_web_tool_patterns.md`
- `reference/12_office_interop_checklist.md`

## 3. 作成したファイル

- なし

## 4. 更新したファイル

- `docs/release_checklist.md`
- `docs/test_scenarios.md`
- `development_report.md`
- `docs/tool_design.md`

## 5. 補完した内容

- リリース前チェックリストを、共通項目と Office Interop 項目に分離した。
- Office 版数・ビット数・利用者権限など、実機で必須確認すべき項目を明文化した。
- テストシナリオに COM 例外系・連続実行時のプロセス残留確認を追加した。
- Markdown 整形品質の自己点検欄を追加した。

## 6. 実装した機能

- 実装機能の追加はなし（既存目的を維持）。
- 文書面での運用・検証観点を補強。
- C#プロジェクト構成を設計書へ明記し、guideで求める構成確認を追記。

## 7. 整形確認結果

- `docs/release_checklist.md` と `docs/test_scenarios.md` の見出し、表、チェックボックス整形を確認。
- 行圧縮はなく、raw表示でも判読可能。

## 8. ビルド確認結果または未確認理由

- `dotnet build src/md2doc/md2doc.csproj` は未実施。
- 理由: 本実行環境に `dotnet` がないため。

## 9. 外部依存の有無

- 既存どおり Microsoft Word（Office Interop）に依存。
- 新規の外部ライブラリ／外部サービス依存は追加なし。

## 10. 外部通信の有無

- なし（オフライン運用前提を維持）。

## 11. 個人情報保存・送信・ログ出力の有無

- 入力本文の永続保存なし。
- 外部送信なし。
- 個人情報を含むログ出力機能の追加なし。

## 12. 判断しづらかった点

- Office 実機未確認時に、どこまでを「判定済み」と扱うかの境界。
- 32bit / 64bit 差異検証の優先度と検証順。

## 13. `lg_toolkit_guide` 側へフィードバックすべき改善点

- Office Interop チェックリストに「実機未確認時の記録テンプレート（未確認理由・次アクション）」を追記すると運用しやすい。
- 非Webツール向けに、WinFormsの典型的な運用引継ぎ記載例があると再利用性が高まる。
