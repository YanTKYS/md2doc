using Md2Doc.Core;
using Xunit;

namespace Md2Doc.Tests;

/// <summary>
/// OpenXmlConverter（Markdig + Open XML SDK）のテスト。
/// Microsoft Word 不要 — すべての環境で実行可能。
/// DocxInspector を共用し、WordInteropConverter テストと同一観点で検証する。
/// </summary>
public class OpenXmlConverterTests : IDisposable
{
    private readonly List<string> _tempFiles = [];

    private string TempDocx()
    {
        var path = Path.Combine(Path.GetTempPath(), $"md2doc_openxml_{Guid.NewGuid():N}.docx");
        _tempFiles.Add(path);
        return path;
    }

    public void Dispose()
    {
        foreach (var f in _tempFiles)
            try { if (File.Exists(f)) File.Delete(f); } catch { }
    }

    private static void Convert(string markdown, string outputPath, bool numberHeadings = false) =>
        OpenXmlConverter.ConvertToDocx(
            markdown, outputPath,
            bodyFontName: "MS Gothic", bodyFontSize: 11.0,
            numberHeadings: numberHeadings,
            headerText: null, headerAlignment: 0,
            addPageNumbers: false, footerAlignment: 1);

    // ──────────────────────────────────────────────────────────────────
    // 要素別テスト（v0.5.1 から継続）
    // ──────────────────────────────────────────────────────────────────

    [Fact]
    public void Headings_H1ToH3_AllPresent()
    {
        const string md = "# H1見出し\n## H2見出し\n### H3見出し";
        var path = TempDocx();
        Convert(md, path);

        var paras = DocxInspector.ExtractParagraphTexts(path);
        var (ok, msg) = DocxInspector.VerifyOrder(paras, ["H1見出し", "H2見出し", "H3見出し"]);
        Assert.True(ok, msg);
    }

    [Fact]
    public void Paragraph_JapaneseText_Preserved()
    {
        const string md = "日本語のテキストが正しく出力される";
        var path = TempDocx();
        Convert(md, path);

        var paras = DocxInspector.ExtractParagraphTexts(path);
        Assert.Contains(paras, p => p.Contains("日本語のテキストが正しく出力される"));
    }

    [Fact]
    public void BulletList_AllItemsPresent()
    {
        const string md = "- 項目1\n- 項目2\n- 項目3\n- 項目4";
        var path = TempDocx();
        Convert(md, path);

        var paras = DocxInspector.ExtractParagraphTexts(path);
        var (ok, msg) = DocxInspector.VerifyOrder(paras, ["項目1", "項目2", "項目3", "項目4"]);
        Assert.True(ok, msg);
    }

    [Fact]
    public void OrderedList_AllItemsPresent()
    {
        const string md = "1. 手順1\n2. 手順2\n3. 手順3";
        var path = TempDocx();
        Convert(md, path);

        var paras = DocxInspector.ExtractParagraphTexts(path);
        var (ok, msg) = DocxInspector.VerifyOrder(paras, ["手順1", "手順2", "手順3"]);
        Assert.True(ok, msg);
    }

    [Fact]
    public void OrderedList_MultipleListsNumberIndependently()
    {
        const string md = "1. A-1\n2. A-2\n\n1. B-1\n2. B-2";
        var path = TempDocx();
        Convert(md, path);

        var paras = DocxInspector.ExtractParagraphTexts(path);
        var (ok, msg) = DocxInspector.VerifyOrder(paras, ["A-1", "A-2", "B-1", "B-2"]);
        Assert.True(ok, msg);
    }

    [Fact]
    public void Table_HeaderAndDataCellsPresent()
    {
        const string md = "| 名前 | 値 |\n|---|---|\n| Alpha | 100 |\n| Beta | 200 |";
        var path = TempDocx();
        Convert(md, path);

        var paras = DocxInspector.ExtractParagraphTexts(path);
        Assert.Contains(paras, p => p.Contains("名前"));
        Assert.Contains(paras, p => p.Contains("値"));
        Assert.Contains(paras, p => p.Contains("Alpha"));
        Assert.Contains(paras, p => p.Contains("100"));
        Assert.Contains(paras, p => p.Contains("Beta"));
        Assert.Contains(paras, p => p.Contains("200"));
    }

    [Fact]
    public void PageBreak_HtmlComment_BothSidesPreserved()
    {
        const string md = "前の段落\n<!-- pagebreak -->\n後の段落";
        var path = TempDocx();
        Convert(md, path);

        var paras = DocxInspector.ExtractParagraphTexts(path);
        var (ok, msg) = DocxInspector.VerifyOrder(paras, ["前の段落", "後の段落"]);
        Assert.True(ok, msg);
    }

    [Fact]
    public void PageBreak_DashSyntax_BothSidesPreserved()
    {
        const string md = "前の段落\n\n---pagebreak---\n\n後の段落";
        var path = TempDocx();
        Convert(md, path);

        var paras = DocxInspector.ExtractParagraphTexts(path);
        var (ok, msg) = DocxInspector.VerifyOrder(paras, ["前の段落", "後の段落"]);
        Assert.True(ok, msg);
    }

    [Fact]
    public void Headings_NumberingEnabled_PrefixAdded()
    {
        const string md = "# 章\n## 節\n### 項";
        var path = TempDocx();
        Convert(md, path, numberHeadings: true);

        var paras = DocxInspector.ExtractParagraphTexts(path);
        Assert.Contains(paras, p => p.Contains("1.") && p.Contains("章"));
        Assert.Contains(paras, p => p.Contains("1.1") && p.Contains("節"));
        Assert.Contains(paras, p => p.Contains("1.1.1") && p.Contains("項"));
    }

    // ──────────────────────────────────────────────────────────────────
    // v0.5.2 新機能テスト
    // ──────────────────────────────────────────────────────────────────

    [Fact]
    public void Header_TextPresent()
    {
        const string md = "段落テキスト";
        var path = TempDocx();
        OpenXmlConverter.ConvertToDocx(md, path,
            bodyFontName: "MS Gothic", bodyFontSize: 11.0,
            numberHeadings: false,
            headerText: "ヘッダーテキスト", headerAlignment: 1,
            addPageNumbers: false, footerAlignment: 1);

        var headerTexts = DocxInspector.ExtractHeaderTexts(path);
        Assert.Contains(headerTexts, t => t.Contains("ヘッダーテキスト"));
    }

    [Fact]
    public void Footer_PageNumber_Present()
    {
        const string md = "段落テキスト";
        var path = TempDocx();
        OpenXmlConverter.ConvertToDocx(md, path,
            bodyFontName: "MS Gothic", bodyFontSize: 11.0,
            numberHeadings: false,
            headerText: null, headerAlignment: 0,
            addPageNumbers: true, footerAlignment: 1);

        Assert.True(DocxInspector.FooterHasPageNumber(path));
    }

    [Fact]
    public void HorizontalRule_Present()
    {
        const string md = "前の段落\n\n---\n\n後の段落";
        var path = TempDocx();
        Convert(md, path);

        Assert.True(DocxInspector.DocumentHasHorizontalRule(path));
        var paras = DocxInspector.ExtractParagraphTexts(path);
        var (ok, msg) = DocxInspector.VerifyOrder(paras, ["前の段落", "後の段落"]);
        Assert.True(ok, msg);
    }

    [Fact]
    public void InlineFormatting_Bold_Present()
    {
        const string md = "通常テキスト **太字テキスト** 通常テキスト";
        var path = TempDocx();
        Convert(md, path);

        Assert.True(DocxInspector.HasRunWithTextAndProperty(path, "太字テキスト", "b"),
            "太字テキストを含む Run に <w:b/> が存在しません。");
    }

    [Fact]
    public void InlineFormatting_Italic_Present()
    {
        const string md = "通常テキスト *斜体テキスト* 通常テキスト";
        var path = TempDocx();
        Convert(md, path);

        Assert.True(DocxInspector.HasRunWithTextAndProperty(path, "斜体テキスト", "i"),
            "斜体テキストを含む Run に <w:i/> が存在しません。");
    }

    [Fact]
    public void InlineFormatting_BoldAndItalicCombined()
    {
        const string md = "***太字斜体テキスト***";
        var path = TempDocx();
        Convert(md, path);

        Assert.True(DocxInspector.HasRunWithTextAndProperty(path, "太字斜体テキスト", "b"),
            "太字斜体テキストに <w:b/> が存在しません。");
        Assert.True(DocxInspector.HasRunWithTextAndProperty(path, "太字斜体テキスト", "i"),
            "太字斜体テキストに <w:i/> が存在しません。");
    }

    [Fact]
    public void InlineFormatting_InlineCode_UsesCodeFont()
    {
        const string md = "通常テキスト `コードテキスト` 通常テキスト";
        var path = TempDocx();
        Convert(md, path);

        Assert.True(DocxInspector.HasRunWithFont(path, "コードテキスト", "Courier New"),
            "インラインコードテキストが Courier New フォントで出力されていません。");
    }

    [Fact]
    public void SoftReturn_HardBreak_Preserved()
    {
        // Markdown の硬改行（行末に2スペース + 改行）→ w:br（ソフトリターン）
        const string md = "行1  \n行2";
        var path = TempDocx();
        Convert(md, path);

        Assert.True(DocxInspector.DocumentHasSoftReturn(path),
            "硬改行が w:br として出力されていません。");
        // 両テキストは同一段落内に存在する（段落分割ではなく行内改行）
        var paras = DocxInspector.ExtractParagraphTexts(path);
        Assert.Contains(paras, p => p.Contains("行1") && p.Contains("行2"));
    }

    [Fact]
    public void SoftReturn_BrTag_Preserved()
    {
        const string md = "行1<br>行2";
        var path = TempDocx();
        Convert(md, path);

        Assert.True(DocxInspector.DocumentHasSoftReturn(path),
            "<br> タグが w:br として出力されていません。");
        var paras = DocxInspector.ExtractParagraphTexts(path);
        Assert.Contains(paras, p => p.Contains("行1") && p.Contains("行2"));
    }

    // ──────────────────────────────────────────────────────────────────
    // 回帰テスト（WordInteropConverter テストと同一 Markdown・同一観点）
    // ──────────────────────────────────────────────────────────────────

    /// <summary>
    /// WordInteropConverter の Regression_AllTextPresentInOrder と同一 Markdown・同一観点。
    /// </summary>
    [Fact]
    public void Regression_AllTextPresentInOrder()
    {
        const string markdown = """
            # タイトル
            ## 起動方法
            - アプリを起動する
            - Markdownファイルを選択する
            本文1
            - 箇条書きA
            - 箇条書きB
            本文2
            """;

        var path = TempDocx();
        Convert(markdown, path);

        var paras = DocxInspector.ExtractParagraphTexts(path);
        string[] expected = [
            "タイトル", "起動方法",
            "アプリを起動する", "Markdownファイルを選択する",
            "本文1", "箇条書きA", "箇条書きB", "本文2",
        ];
        var (ok, msg) = DocxInspector.VerifyOrder(paras, expected);
        Assert.True(ok, msg);
    }

    [Fact]
    public void Regression_NoBulletAlternatingLoss()
    {
        const string markdown = """
            ## 見出し
            - 1行目
            - 2行目
            - 3行目
            - 4行目
            """;

        var path = TempDocx();
        Convert(markdown, path);

        var paras = DocxInspector.ExtractParagraphTexts(path);
        var (ok, msg) = DocxInspector.VerifyOrder(paras, ["見出し", "1行目", "2行目", "3行目", "4行目"]);
        Assert.True(ok, msg);
    }

    [Fact]
    public void Regression_OrderedAndUnorderedMixed()
    {
        const string markdown = """
            ## 起動方法
            - アプリを起動する
            - Markdownファイルを選択する

            ## 変換方法
            1. 出力先を確認する
            2. 変換実行を押す
            """;

        var path = TempDocx();
        Convert(markdown, path);

        var paras = DocxInspector.ExtractParagraphTexts(path);
        var (ok, msg) = DocxInspector.VerifyOrder(paras, [
            "起動方法",
            "アプリを起動する", "Markdownファイルを選択する",
            "変換方法",
            "出力先を確認する", "変換実行を押す",
        ]);
        Assert.True(ok, msg);
    }

    [Fact]
    public void Regression_BulletsAroundTable()
    {
        const string markdown = """
            ## 表の前
            - 箇条書き1
            - 箇条書き2

            | A | B |
            |---|---|
            | 1 | 2 |

            - 表の後の箇条書き
            本文
            """;

        var path = TempDocx();
        Convert(markdown, path);

        var paras = DocxInspector.ExtractParagraphTexts(path);
        var (ok, msg) = DocxInspector.VerifyOrder(paras,
            ["表の前", "箇条書き1", "箇条書き2", "表の後の箇条書き", "本文"]);
        Assert.True(ok, msg);
    }
}
