# 11_non_web_tool_patterns.md - 非Webツール標準構成パターン

## 1. 目的

- `lg_toolkit_guide` が静的Webツールだけでなく、PowerShellやC#ツールにも適用できることを明確にする。
- 実装方式ごとの標準ディレクトリ構成を分ける。
- Web標準構成とC#標準構成が混在して見えないようにする。

## 2. 実装方式別の標準構成

### 2.1 静的Webツール

```text
src/
  index.html
  script.js
  style.css
```

### 2.2 PowerShellツール

```text
src/
  main.ps1
```

必要に応じて:

```text
src/
  modules/
```

### 2.3 C# Consoleツール

```text
src/
  <AppName>/
    <AppName>.csproj
    Program.cs
```

### 2.4 C# WinFormsツール

```text
src/
  <AppName>/
    <AppName>.csproj
    Program.cs
    MainForm.cs
    （必要に応じてその他クラス）
```

### 2.5 Office Interop系C#ツール

```text
src/
  <AppName>/
    <AppName>.csproj
    Program.cs
    MainForm.cs
    <OfficeInteropService>.cs
```

## 3. 共通成果物

実装方式が変わっても、原則として以下の文書成果物は共通とする。

```text
README.md
development_report.md

docs/
  tool_design.md
  release_checklist.md
  test_scenarios.md
  operation_handover.md

manuals/
  admin_manual.md
  operator_manual.md
  user_manual.md

reference/
  guide_context.md
```

- `reference/guide_context.md` は同梱方式の場合のみ必要。

## 4. 注意点

- `src/index.html` 等は静的Webツール向け標準構成であり、すべてのツールに必須ではない。
- C#やPowerShellでは、実装方式に応じて `src/` 配下の構成を変える。
- 文書成果物は共通、実装成果物は方式別と整理する。
