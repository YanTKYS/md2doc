using System.Text.RegularExpressions;

namespace Md2Doc;

internal static class WordInteropConverter
{
    private static readonly Regex HeadingPattern = new(@"^(#{1,6})\s+(.+)$", RegexOptions.Compiled);
    private static readonly Regex BulletPattern = new(@"^[-*]\s+(.+)$", RegexOptions.Compiled);
    private static readonly Regex BoldPattern = new(@"\*\*(.+?)\*\*", RegexOptions.Compiled);
    private static readonly Regex ItalicPattern = new(@"\*(.+?)\*", RegexOptions.Compiled);
    private static readonly Regex CodePattern = new(@"`(.+?)`", RegexOptions.Compiled);

    public static void ConvertToDocx(
        string markdown,
        string outputPath,
        string headingFontName,
        double headingFontSize,
        string bodyFontName,
        double bodyFontSize,
        string? headerText,
        int headerAlignment,
        bool addPageNumbers,
        int footerAlignment)
    {
        var wordType = Type.GetTypeFromProgID("Word.Application")
            ?? throw new InvalidOperationException("Microsoft Word が利用できません。");

        dynamic? app = null;
        dynamic? doc = null;

        try
        {
            AppLog.Info($"変換開始: output={outputPath} headingFont={headingFontName} bodyFont={bodyFontName}");
            app = Activator.CreateInstance(wordType)
                ?? throw new InvalidOperationException("Wordを起動できません。");
            app.Visible = false;
            doc = app.Documents.Add();

            SetAuthor(doc);
            WriteMarkdown(doc, markdown, headingFontName, headingFontSize, bodyFontName, bodyFontSize);

            if (headerText is not null)
                SetHeader(doc, headerText, headerAlignment);

            if (addPageNumbers)
                SetFooterPageNumbers(doc, footerAlignment);

            doc.SaveAs2(outputPath, 12); // 12 = wdFormatXMLDocument (.docx)
            AppLog.Info($"変換完了: {outputPath}");
        }
        finally
        {
            // Close/Quit を個別の try-catch で保護し、例外が出ても必ず FinalReleaseComObject を実行する
            if (doc is not null)
            {
                try { doc.Close(false); }
                catch (Exception ex) { AppLog.Error("doc.Close 失敗", ex); }
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(doc);
            }
            if (app is not null)
            {
                try { app.Quit(false); }
                catch (Exception ex) { AppLog.Error("app.Quit 失敗", ex); }
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(app);
            }
        }
    }

    private static void SetAuthor(dynamic doc)
    {
        try
        {
            doc.BuiltInDocumentProperties("Author").Value = "md2doc";
        }
        catch
        {
            // 作者名の設定失敗は変換結果に影響しない
        }
    }

    private static void WriteMarkdown(
        dynamic doc,
        string markdown,
        string headingFontName,
        double headingFontSize,
        string bodyFontName,
        double bodyFontSize)
    {
        var lines = markdown.Replace("\r\n", "\n").Split('\n');
        bool firstParagraphUsed = false;

        foreach (var rawLine in lines)
        {
            var line = rawLine.TrimEnd();
            if (string.IsNullOrWhiteSpace(line))
            {
                if (firstParagraphUsed)
                    doc.Paragraphs.Add();
                continue;
            }

            // 新規ドキュメントの最初の空段落を再利用することで先頭の余分な空行を防ぐ
            dynamic para = firstParagraphUsed ? doc.Paragraphs.Add() : doc.Paragraphs[1];
            firstParagraphUsed = true;

            if (!TryHeading(para, line, headingFontName, headingFontSize) &&
                !TryBullet(para, line, bodyFontName, bodyFontSize))
            {
                WriteParagraph(para, ParseInline(line), bodyFontName, bodyFontSize);
            }
        }
    }

    private static bool TryHeading(dynamic para, string line, string fontName, double fontSize)
    {
        var match = HeadingPattern.Match(line);
        if (!match.Success)
            return false;

        var level = Math.Min(match.Groups[1].Value.Length, 3);
        para.Range.Text = ParseInline(match.Groups[2].Value);
        // WdBuiltinStyle 定数を使用（言語非依存: Heading N = -(N+1)）
        para.Range.Style = -(level + 1);
        // スタイル適用後にフォントを上書き（直接書式はスタイル書式より優先される）
        para.Range.Font.Name = fontName;
        para.Range.Font.Size = fontSize;
        return true;
    }

    private static bool TryBullet(dynamic para, string line, string fontName, double fontSize)
    {
        var match = BulletPattern.Match(line);
        if (!match.Success)
            return false;

        para.Range.Text = ParseInline(match.Groups[1].Value);
        para.Range.ListFormat.ApplyBulletDefault();
        para.Range.Font.Name = fontName;
        para.Range.Font.Size = fontSize;
        return true;
    }

    private static void WriteParagraph(dynamic para, string text, string fontName, double fontSize)
    {
        para.Range.Text = text;
        para.Range.Font.Name = fontName;
        para.Range.Font.Size = fontSize;
    }

    private static void SetHeader(dynamic doc, string headerText, int alignment)
    {
        // 1 = wdHeaderFooterPrimary
        dynamic header = doc.Sections[1].Headers[1];
        header.Range.Text = headerText;
        header.Range.ParagraphFormat.Alignment = alignment;
    }

    private static void SetFooterPageNumbers(dynamic doc, int alignment)
    {
        // 1 = wdHeaderFooterPrimary
        dynamic footer = doc.Sections[1].Footers[1];
        dynamic range = footer.Range;

        // PAGE フィールドを挿入（33 = wdFieldPage）
        range.Fields.Add(range, 33);

        // フィールド挿入後に Range を再取得してセパレータを追記
        range = footer.Range;
        range.InsertAfter(" / ");
        range.Collapse(0); // 0 = wdCollapseEnd

        // NUMPAGES フィールドを挿入（26 = wdFieldNumPages）
        range.Fields.Add(range, 26);

        footer.Range.ParagraphFormat.Alignment = alignment;
    }

    private static string ParseInline(string text)
    {
        var result = BoldPattern.Replace(text, "$1");
        result = ItalicPattern.Replace(result, "$1");
        result = CodePattern.Replace(result, "$1");
        return result;
    }
}
