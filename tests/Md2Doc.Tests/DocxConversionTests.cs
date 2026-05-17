using System.Runtime.InteropServices;
using Xunit;

namespace Md2Doc.Tests;

/// <summary>
/// Word COM で .docx を生成し、word/document.xml を検査する統合テスト。
/// Windows + Microsoft Word がインストールされていない環境では自動スキップする。
/// </summary>
public class DocxConversionTests : IDisposable
{
    private readonly List<string> _tempFiles = [];

    private static bool WordAvailable() =>
        RuntimeInformation.IsOSPlatform(OSPlatform.Windows) &&
        Type.GetTypeFromProgID("Word.Application") != null;

    private string TempDocx()
    {
        var path = Path.Combine(Path.GetTempPath(), $"md2doc_test_{Guid.NewGuid():N}.docx");
        _tempFiles.Add(path);
        return path;
    }

    public void Dispose()
    {
        foreach (var f in _tempFiles)
            try { if (File.Exists(f)) File.Delete(f); } catch { }
    }

    // -----------------------------------------------------------------------
    // 必須テストケース: 回帰テスト仕様 Markdown
    // -----------------------------------------------------------------------

    /// <summary>
    /// 全テキストが document.xml に存在し、Markdown と同じ順序で並んでいることを確認する。
    /// 見出し直後の箇条書き 1 行目消失、箇条書きの交互消失、
    /// 箇条書き後の通常段落消失 を一括で検出できる。
    /// </summary>
    [Fact]
    public void Regression_AllTextPresentInOrder()
    {
        if (!WordAvailable()) return;

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

        var output = TempDocx();
        WordInteropConverter.ConvertToDocx(
            markdown, output,
            bodyFontName: "MS Gothic", bodyFontSize: 11.0,
            numberHeadings: false,
            headerText: null, headerAlignment: 0,
            addPageNumbers: false, footerAlignment: 1);

        var paragraphs = DocxInspector.ExtractParagraphTexts(output);

        string[] expected = [
            "タイトル",
            "起動方法",
            "アプリを起動する",
            "Markdownファイルを選択する",
            "本文1",
            "箇条書きA",
            "箇条書きB",
            "本文2",
        ];

        var (ok, msg) = DocxInspector.VerifyOrder(paragraphs, expected);
        Assert.True(ok, msg);
    }

    /// <summary>
    /// 箇条書きが奇数行または偶数行だけに減っていないことを確認する。
    /// 交互消失パターン（v0.3.0 で発生した問題）の回帰テスト。
    /// </summary>
    [Fact]
    public void Regression_NoBulletAlternatingLoss()
    {
        if (!WordAvailable()) return;

        const string markdown = """
            ## 見出し
            - 1行目
            - 2行目
            - 3行目
            - 4行目
            """;

        var output = TempDocx();
        WordInteropConverter.ConvertToDocx(
            markdown, output,
            bodyFontName: "MS Gothic", bodyFontSize: 11.0,
            numberHeadings: false,
            headerText: null, headerAlignment: 0,
            addPageNumbers: false, footerAlignment: 1);

        var paragraphs = DocxInspector.ExtractParagraphTexts(output);

        var (ok, msg) = DocxInspector.VerifyOrder(paragraphs, ["見出し", "1行目", "2行目", "3行目", "4行目"]);
        Assert.True(ok, msg);
    }

    /// <summary>
    /// 箇条書き後の通常段落にリスト書式が伝播していないことを確認する。
    /// 段落がリスト化されると段落テキストは存在するが書式が崩れる。
    /// </summary>
    [Fact]
    public void Regression_ParagraphAfterBulletNotLost()
    {
        if (!WordAvailable()) return;

        const string markdown = """
            - 箇条書き1
            - 箇条書き2
            通常段落
            """;

        var output = TempDocx();
        WordInteropConverter.ConvertToDocx(
            markdown, output,
            bodyFontName: "MS Gothic", bodyFontSize: 11.0,
            numberHeadings: false,
            headerText: null, headerAlignment: 0,
            addPageNumbers: false, footerAlignment: 1);

        var paragraphs = DocxInspector.ExtractParagraphTexts(output);

        var (ok, msg) = DocxInspector.VerifyOrder(paragraphs, ["箇条書き1", "箇条書き2", "通常段落"]);
        Assert.True(ok, msg);
    }

    // -----------------------------------------------------------------------
    // 追加テストケース: テーブル前後の箇条書き
    // -----------------------------------------------------------------------

    /// <summary>
    /// テーブル前後の箇条書きが消えないことを確認する。
    /// テーブル挿入（Phase 4）によるインデックスシフト・リスト伝播の複合ケース。
    /// </summary>
    [Fact]
    public void Regression_BulletsAroundTable()
    {
        if (!WordAvailable()) return;

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

        var output = TempDocx();
        WordInteropConverter.ConvertToDocx(
            markdown, output,
            bodyFontName: "MS Gothic", bodyFontSize: 11.0,
            numberHeadings: false,
            headerText: null, headerAlignment: 0,
            addPageNumbers: false, footerAlignment: 1);

        var paragraphs = DocxInspector.ExtractParagraphTexts(output);

        var (ok, msg) = DocxInspector.VerifyOrder(paragraphs,
            ["表の前", "箇条書き1", "箇条書き2", "表の後の箇条書き", "本文"]);
        Assert.True(ok, msg);
    }

    /// <summary>
    /// 見出し番号付与オプションが有効なとき、番号付き見出しが出力されることを確認する。
    /// </summary>
    [Fact]
    public void Regression_NumberedHeadings_TextPresent()
    {
        if (!WordAvailable()) return;

        const string markdown = """
            # 第1章
            ## 節1
            - 箇条書き
            本文
            """;

        var output = TempDocx();
        WordInteropConverter.ConvertToDocx(
            markdown, output,
            bodyFontName: "MS Gothic", bodyFontSize: 11.0,
            numberHeadings: true,
            headerText: null, headerAlignment: 0,
            addPageNumbers: false, footerAlignment: 1);

        var paragraphs = DocxInspector.ExtractParagraphTexts(output);

        // 番号付き見出しは "1. 第1章" / "1.1 節1" になる
        var (ok, msg) = DocxInspector.VerifyOrder(paragraphs, ["第1章", "節1", "箇条書き", "本文"]);
        Assert.True(ok, msg);
    }
}
