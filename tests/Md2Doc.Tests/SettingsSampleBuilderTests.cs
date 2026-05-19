using Md2Doc.Core;
using Xunit;

namespace Md2Doc.Tests;

public class SettingsSampleBuilderTests
{
    // ──────────────────────────────────────────────────────
    // ヘッダーなし
    // ──────────────────────────────────────────────────────

    [Fact]
    public void Header_None_ShowsNashi()
    {
        var result = Build(headerMode: 0);
        Assert.Contains("【ヘッダー】", result);
        Assert.Contains("（なし）", result);
        Assert.DoesNotContain("左：", result);
    }

    // ──────────────────────────────────────────────────────
    // ヘッダー自由記入 × 配置
    // ──────────────────────────────────────────────────────

    [Fact]
    public void Header_Custom_Left_PlacedCorrectly()
    {
        var result = Build(headerMode: 2, headerCustomText: "業務手順書", headerAlignment: 0);
        Assert.Contains("左：業務手順書", result);
        Assert.Contains("中央：（なし）", result);
        Assert.Contains("右：（なし）", result);
    }

    [Fact]
    public void Header_Custom_Center_PlacedCorrectly()
    {
        var result = Build(headerMode: 2, headerCustomText: "業務手順書", headerAlignment: 1);
        Assert.Contains("左：（なし）", result);
        Assert.Contains("中央：業務手順書", result);
        Assert.Contains("右：（なし）", result);
    }

    [Fact]
    public void Header_Custom_Right_PlacedCorrectly()
    {
        var result = Build(headerMode: 2, headerCustomText: "業務手順書", headerAlignment: 2);
        Assert.Contains("左：（なし）", result);
        Assert.Contains("中央：（なし）", result);
        Assert.Contains("右：業務手順書", result);
    }

    [Fact]
    public void Header_Custom_EmptyText_ShowsUnnyuryoku()
    {
        var result = Build(headerMode: 2, headerCustomText: "");
        Assert.Contains("（未入力）", result);
    }

    // ──────────────────────────────────────────────────────
    // ヘッダー ファイル名
    // ──────────────────────────────────────────────────────

    [Fact]
    public void Header_FileName_WithName_ShowsName()
    {
        var result = Build(headerMode: 1, inputFileName: "report.md", headerAlignment: 0);
        Assert.Contains("左：report.md", result);
    }

    [Fact]
    public void Header_FileName_NoFile_ShowsPlaceholder()
    {
        var result = Build(headerMode: 1, inputFileName: null);
        Assert.Contains("（ファイル名）", result);
    }

    [Fact]
    public void Header_FileName_EmptyString_ShowsPlaceholder()
    {
        var result = Build(headerMode: 1, inputFileName: "");
        Assert.Contains("（ファイル名）", result);
    }

    // ──────────────────────────────────────────────────────
    // 見出し番号
    // ──────────────────────────────────────────────────────

    [Fact]
    public void NumberHeadings_Off_ShowsNashi()
    {
        var result = Build(numberHeadings: false);
        var headingSection = result.Substring(result.IndexOf("【見出し番号】"));
        Assert.Contains("（なし）", headingSection);
        Assert.DoesNotContain("1. 大見出し", result);
    }

    [Fact]
    public void NumberHeadings_On_ShowsSample()
    {
        var result = Build(numberHeadings: true);
        Assert.Contains("1. 大見出し", result);
        Assert.Contains("1.1 中見出し", result);
        Assert.Contains("1.1.1 小見出し", result);
    }

    // ──────────────────────────────────────────────────────
    // ページ番号
    // ──────────────────────────────────────────────────────

    [Fact]
    public void FooterPageNumber_Off_ShowsNashi()
    {
        var result = Build(footerPageNumber: false);
        var footerSection = result.Substring(result.IndexOf("【ページ番号】"));
        Assert.Contains("（なし）", footerSection);
        Assert.DoesNotContain("1 / 5", result);
    }

    [Fact]
    public void FooterPageNumber_On_ShowsSample()
    {
        var result = Build(footerPageNumber: true);
        Assert.Contains("1 / 5", result);
    }

    // ──────────────────────────────────────────────────────
    // セクション見出しの存在確認
    // ──────────────────────────────────────────────────────

    [Fact]
    public void AllSections_AlwaysPresent()
    {
        var result = Build();
        Assert.Contains("【ヘッダー】", result);
        Assert.Contains("【見出し番号】", result);
        Assert.Contains("【ページ番号】", result);
    }

    // ──────────────────────────────────────────────────────
    // ヘルパー
    // ──────────────────────────────────────────────────────

    private static string Build(
        int headerMode = 0,
        string headerCustomText = "",
        string? inputFileName = null,
        int headerAlignment = 0,
        bool numberHeadings = false,
        bool footerPageNumber = false) =>
        SettingsSampleBuilder.Build(
            headerMode, headerCustomText, inputFileName,
            headerAlignment, numberHeadings, footerPageNumber);
}
