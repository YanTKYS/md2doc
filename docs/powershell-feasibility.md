# PowerShell 版実装の実現可能性検討

## 検討概要

「Markdown テキストまたはそのファイルをテキスト入力もしくは D&D で GUI に入力し、
Open XML 方式で Word ファイルを出力する」という現行機能を、
PowerShell で実装可能かどうかを検討した。

実装は行わず、判断のみを目的とする。

---

## 方式の整理

### 方式A：既存 Md2Doc.Core.dll を読み込んで呼び出す

PowerShell は `Add-Type -Path "Md2Doc.Core.dll"` で .NET アセンブリを直接ロードできる。
変換処理は Core DLL に委譲し、PowerShell 側は GUI とファイル操作のグルーコードを担う。

- **Open XML 変換**：`Md2Doc.Core.dll` + `Markdig.dll` + `DocumentFormat.OpenXml.dll` を同梱することで変換ロジックの再実装は不要
- **WinForms GUI**：`[System.Windows.Forms.Form]` 等を用いて構築可能
- **D&D**：`$control.AllowDrop = $true` ＋ `Add_DragEnter` / `Add_DragDrop` イベントで対応可能
- **配布形態**：`.ps1` スクリプト ＋ DLL 一式のフォルダ配布

### 方式B：PowerShell のみで変換ロジックを再実装する

現行の `OpenXmlConverter.cs` は Markdig AST を走査する複雑な C# コード（約 600 行）であり、
PowerShell スクリプトでの再実装は技術的には不可能ではないが、保守コストが著しく高く実用に適しない。
**方式B は非現実的と判断した。**

---

## 方式A の技術的実現性

| 観点 | 評価 |
|------|------|
| 技術的可否 | ✅ 可能 |
| 変換ロジック再実装 | 不要（Core DLL をそのまま使用） |
| WinForms GUI 構築 | ✅ 可能（C# より記述量は多くなる） |
| D&D 対応 | ✅ WinForms イベントで対応可能 |
| Microsoft Word 不要 | ✅ 現行と同様 |

---

## ランタイム同梱に関する制約

現行プロダクトは以下のビルドオプションで .NET 8 ランタイムを exe に内包している。

```
dotnet publish --self-contained true -p:PublishSingleFile=true
```

PowerShell スクリプト（`.ps1`）にはこれに相当する仕組みが存在しない。
`Md2Doc.Core.dll` は `net8.0` 向けビルドであるため、実行環境として以下のいずれかが必要になる。

| 選択肢 | 問題点 |
|--------|--------|
| Windows PowerShell 5.1（OS 標準） | .NET Framework 4.x ベースのため `net8.0` DLL のロード不可 |
| PowerShell 7.x | .NET 8 ベースで動作可能だが、Windows への標準搭載なし。別途インストールが必要 |
| .NET 8 Desktop Runtime 単体インストール | 「ランタイムを別途インストールする」制約は現行より後退 |
| PowerShell 7 をフォルダ同梱 | zip 展開形式での配布は可能だが展開後 100 MB 超となり非現実的 |

### 結論

**PowerShell 方式では、現行と同等の「ランタイム同梱・前提条件ゼロ」は実現できない。**

現行プロダクトの価値のひとつは「利用者端末に追加インストール不要で動作すること」であり、
方式A はその価値を損なう。庁内端末への追加インストールが必要になる点で運用負担が増加する。

方式A が現実的に成立するとすれば、「対象端末に PowerShell 7 がすでに組織として展開済みである」
という前提が確立している場合に限られる。

---

## 総合評価

| 観点 | 評価 |
|------|------|
| 技術的実現性（方式A） | ✅ 可能 |
| ランタイム同梱の再現 | ❌ 不可 |
| 現行と同等の配布容易性 | ❌ 後退する |
| 推奨 | PowerShell 7 が組織展開済みの場合のみ検討対象 |

---

## 関連

- `docs/backlog.md`：PowerShell 版を v1.1.0 以降の検討事項として記録済み
- `src/Md2Doc.Core/`：再利用対象の変換ライブラリ（`net8.0`）
