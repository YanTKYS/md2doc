#Requires -Version 5.1
# Run-IntegrationTests.ps1
# md2doc 統合テスト実行スクリプト
#
# 実行要件:
#   - Windows
#   - Microsoft Word（インストール済み・起動可能な状態）
#   - .NET 8 SDK  ※ Runtime のみでは不可
#     https://dotnet.microsoft.com/download/dotnet/8.0
#
# 使い方:
#   PowerShell を開き、このスクリプトがあるフォルダで実行してください。
#   .\Run-IntegrationTests.ps1

$ErrorActionPreference = 'Stop'
$ScriptDir  = $PSScriptRoot
$TestDll    = Join-Path $ScriptDir "tests\Md2Doc.Tests.dll"

# ────────────────────────────────────────────────
# 前提チェック
# ────────────────────────────────────────────────
Write-Host ""
Write-Host "===  md2doc 統合テスト  ===" -ForegroundColor Cyan
Write-Host ""

# .NET SDK チェック
$dotnetOk = $false
try {
    $ver = & dotnet --version 2>$null
    if ($ver -match '^8\.') {
        Write-Host "[OK] .NET SDK: $ver" -ForegroundColor Green
        $dotnetOk = $true
    } else {
        Write-Host "[NG] .NET 8 SDK が必要です（現在: $ver）" -ForegroundColor Red
        Write-Host "     https://dotnet.microsoft.com/download/dotnet/8.0" -ForegroundColor Yellow
    }
} catch {
    Write-Host "[NG] dotnet コマンドが見つかりません。.NET 8 SDK をインストールしてください。" -ForegroundColor Red
    Write-Host "     https://dotnet.microsoft.com/download/dotnet/8.0" -ForegroundColor Yellow
}

# Microsoft Word チェック
$wordOk = $false
$wordType = [Type]::GetTypeFromProgID("Word.Application")
if ($null -ne $wordType) {
    Write-Host "[OK] Microsoft Word: 検出" -ForegroundColor Green
    $wordOk = $true
} else {
    Write-Host "[NG] Microsoft Word が見つかりません。" -ForegroundColor Red
    Write-Host "     Word がインストールされていないと統合テストは実行できません。" -ForegroundColor Yellow
}

# テスト DLL チェック
$dllOk = Test-Path $TestDll
if ($dllOk) {
    Write-Host "[OK] テスト DLL: $TestDll" -ForegroundColor Green
} else {
    Write-Host "[NG] テスト DLL が見つかりません: $TestDll" -ForegroundColor Red
}

Write-Host ""

if (-not ($dotnetOk -and $wordOk -and $dllOk)) {
    Write-Host "前提条件が満たされていないため、テストを実行できません。" -ForegroundColor Red
    Write-Host "上記の [NG] 項目を解消してから再実行してください。" -ForegroundColor Red
    exit 1
}

# ────────────────────────────────────────────────
# テスト実行
# ────────────────────────────────────────────────
Write-Host "統合テストを実行します..." -ForegroundColor Cyan
Write-Host "（Word が起動します。完了まで数分かかる場合があります）"
Write-Host ""

& dotnet test $TestDll --logger "console;verbosity=normal"
$exitCode = $LASTEXITCODE

Write-Host ""
if ($exitCode -eq 0) {
    Write-Host "すべてのテストが通過しました。" -ForegroundColor Green
} else {
    Write-Host "テストが失敗しました（終了コード: $exitCode）。上記の出力を確認してください。" -ForegroundColor Red
}

exit $exitCode
