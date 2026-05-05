# 開発報告書

## 1. 今回の目的

更新後の `reference/guide_context.md` に合わせて、
既存のWinForms/Office Interop実装を活かしつつ、
不足成果物・運用文書・確認記録を補完する。

## 2. 作成したファイル

- `README.md`
- `docs/release_checklist.md`
- `docs/test_scenarios.md`
- `docs/operation_handover.md`
- `manuals/admin_manual.md`
- `manuals/operator_manual.md`
- `manuals/user_manual.md`

## 3. 更新したファイル

- `src/md2doc/WordInteropConverter.cs`

## 4. 補完した内容

- 標準成果物名に合わせた文書群を追加した。
- WinForms / Office Interop運用時の注意事項をREADMEと運用文書へ記載した。
- テストシナリオとリリースチェックリストを追加した。

## 5. 実装・整形・ビルド確認

- Markdown文書群の改行・見出し・表を整形した。
- C#実装の不要usingを整理した。
- `dotnet build` はこの環境で未実施（`dotnet` 非搭載）。

## 6. 判断しづらかった点

- Office StandardでのWord連携を、
  どの粒度まで標準動作保証とするかは運用定義が必要。

## 7. 未対応事項と対応予定

- Windows端末での実機ビルド確認。
- Office Standard端末での変換実行確認。

## 8. 今回あえて作成しなかったもの

- `src/index.html` `src/script.js` `src/style.css`
  - 本ツールはWinForms実装のため、Web資材は対象外とした。
