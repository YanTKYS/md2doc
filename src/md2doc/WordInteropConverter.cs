using System.Text.RegularExpressions;

namespace Md2Doc;

internal static class WordInteropConverter
{
    private static readonly Regex HeadingPattern = new(@"^(#{1,6})\s+(.+)$", RegexOptions.Compiled);
    private static readonly Regex BulletPattern = new(@"^[-*]\s+(.+)$", RegexOptions.Compiled);
    private static readonly Regex BoldPattern = new(@"\*\*(.+?)\*\*", RegexOptions.Compiled);
    private static readonly Regex ItalicPattern = new(@"\*(.+?)\*", RegexOptions.Compiled);
    private static readonly Regex CodePattern = new(@"`(.+?)`", RegexOptions.Compiled);
    private static readonly Regex BrTagPattern = new(@"<br\s*/?>", RegexOptions.Compiled | RegexOptions.IgnoreCase);
    private static readonly Regex HrPattern = new(@"^[-*_]{3,}\s*$", RegexOptions.Compiled);
    private static readonly Regex PageBreakPattern = new(
        @"^(<!--\s*pagebreak\s*-->|---pagebreak---)$", RegexOptions.Compiled | RegexOptions.IgnoreCase);

    public static void ConvertToDocx(
        string markdown,
        string outputPath,
        string bodyFontName,
        double bodyFontSize,
        bool numberHeadings,
        string? headerText,
        int headerAlignment,
        bool addPageNumbers,
        int footerAlignment,
        IProgress<int>? progress = null)
    {
        var wordType = Type.GetTypeFromProgID("Word.Application")
            ?? throw new InvalidOperationException("Microsoft Word が利用できません。");

        dynamic? app = null;
        dynamic? doc = null;

        try
        {
            AppLog.Info($"変換開始: output={outputPath} bodyFont={bodyFontName} numberHeadings={numberHeadings}");
            progress?.Report(5);

            app = Activator.CreateInstance(wordType)
                ?? throw new InvalidOperationException("Wordを起動できません。");
            app.Visible = false;
            doc = app.Documents.Add();
            progress?.Report(10);

            SetAuthor(doc);
            WriteMarkdown(doc, markdown, bodyFontName, bodyFontSize, numberHeadings, progress);
            progress?.Report(80);

            if (headerText is not null)
                SetHeader(doc, headerText, headerAlignment);
            progress?.Report(85);

            if (addPageNumbers)
                SetFooterPageNumbers(doc, footerAlignment);
            progress?.Report(90);

            doc.SaveAs2(outputPath, 12); // 12 = wdFormatXMLDocument (.docx)
            progress?.Report(95);
            AppLog.Info($"変換完了: {outputPath}");
        }
        finally
        {
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
        catch { }
    }

    private static void WriteMarkdown(
        dynamic doc,
        string markdown,
        string bodyFontName,
        double bodyFontSize,
        bool numberHeadings,
        IProgress<int>? progress)
    {
        var lines = markdown.Replace("\r\n", "\n").Split('\n');
        int total = Math.Max(lines.Length, 1);
        bool firstParagraphUsed = false;
        // 見出し番号カウンター [H1, H2, H3]
        var headingCounters = new int[3];
        int i = 0;

        while (i < lines.Length)
        {
            var line = lines[i].TrimEnd();

            // テーブルブロックの検出：連続する | 行をまとめて処理
            if (IsTableLine(line))
            {
                var tableLines = new List<string>();
                while (i < lines.Length && IsTableLine(lines[i].TrimEnd()))
                {
                    tableLines.Add(lines[i].TrimEnd());
                    i++;
                }
                if (TryTable(doc, tableLines, bodyFontName, bodyFontSize, firstParagraphUsed))
                    firstParagraphUsed = true;
                progress?.Report(10 + i * 70 / total);
                continue;
            }

            if (string.IsNullOrWhiteSpace(line))
            {
                if (firstParagraphUsed)
                    doc.Paragraphs.Add();
            }
            else
            {
                // 新規ドキュメントの最初の空段落を再利用することで先頭の余分な空行を防ぐ
                dynamic para = firstParagraphUsed ? doc.Paragraphs.Add() : doc.Paragraphs[1];
                firstParagraphUsed = true;

                if (!TryHeading(para, line, numberHeadings, headingCounters) &&
                    !TryBullet(para, line, bodyFontName, bodyFontSize) &&
                    !TryHorizontalRule(para, line) &&
                    !TryPageBreak(para, line))
                {
                    WriteParagraph(para, ParseInline(line), bodyFontName, bodyFontSize);
                }
            }

            // 10%→80% の範囲で行ごとに進捗報告
            progress?.Report(10 + (i + 1) * 70 / total);
            i++;
        }
    }

    private static bool IsTableLine(string line)
    {
        var t = line.Trim();
        return t.Length > 2 && t.StartsWith("|") && t.EndsWith("|");
    }

    private static bool IsTableSeparator(string line) =>
        line.All(c => c is '-' or ':' or '|' or ' ');

    private static List<string> ParseTableRow(string line)
    {
        var parts = line.Split('|');
        return parts.Length < 2 ? [] : parts[1..^1].Select(p => p.Trim()).ToList();
    }

    private static bool TryTable(
        dynamic doc, List<string> tableLines,
        string fontName, double fontSize, bool firstParagraphUsed)
    {
        bool hasSeparator = tableLines.Any(IsTableSeparator);
        var dataRows = tableLines.Where(l => !IsTableSeparator(l)).ToList();
        if (dataRows.Count == 0) return false;

        var colCount = ParseTableRow(dataRows[0]).Count;
        if (colCount == 0) return false;

        dynamic anchorPara = firstParagraphUsed ? doc.Paragraphs.Add() : doc.Paragraphs[1];
        dynamic table = doc.Tables.Add(anchorPara.Range, dataRows.Count, colCount);
        table.Borders.Enable = 1;

        for (int r = 0; r < dataRows.Count; r++)
        {
            var cells = ParseTableRow(dataRows[r]);
            for (int c = 0; c < Math.Min(cells.Count, colCount); c++)
            {
                var cell = table.Cell(r + 1, c + 1);
                cell.Range.Text = ParseInline(cells[c]);
                cell.Range.Font.Name = fontName;
                cell.Range.Font.Size = fontSize;
                // 区切り行がある場合は先頭行をヘッダーとして太字にする
                if (hasSeparator && r == 0)
                    cell.Range.Font.Bold = 1;
            }
        }

        return true;
    }

    private static bool TryHeading(dynamic para, string line, bool numberHeadings, int[] counters)
    {
        var match = HeadingPattern.Match(line);
        if (!match.Success) return false;

        var level = Math.Min(match.Groups[1].Value.Length, 3);
        var text = ParseInline(match.Groups[2].Value);

        if (numberHeadings)
        {
            counters[level - 1]++;
            for (int i = level; i < 3; i++) counters[i] = 0;

            var prefix = level switch
            {
                1 => $"{counters[0]}. ",
                2 => $"{counters[0]}.{counters[1]} ",
                _ => $"{counters[0]}.{counters[1]}.{counters[2]} ",
            };
            text = prefix + text;
        }

        para.Range.Text = text;
        // WdBuiltinStyle 定数を使用（言語非依存: Heading N = -(N+1)）
        // フォント上書きなし — Wordの標準見出しスタイルの書式をそのまま適用する
        para.Range.Style = -(level + 1);
        return true;
    }

    private static bool TryHorizontalRule(dynamic para, string line)
    {
        if (!HrPattern.IsMatch(line)) return false;
        // -3 = wdBorderBottom, 1 = wdLineStyleSingle
        para.Borders[-3].LineStyle = 1;
        return true;
    }

    private static bool TryPageBreak(dynamic para, string line)
    {
        if (!PageBreakPattern.IsMatch(line)) return false;
        // \f = Chr(12) = wdPageBreak — 改ページ文字を段落テキストとして挿入
        para.Range.Text = "\f";
        return true;
    }

    private static bool TryBullet(dynamic para, string line, string fontName, double fontSize)
    {
        var match = BulletPattern.Match(line);
        if (!match.Success) return false;

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
        // <br> / <br/> / <BR> 等を Word のソフトリターン（Chr(11)）に変換
        var result = BrTagPattern.Replace(text, "\v");
        result = BoldPattern.Replace(result, "$1");
        result = ItalicPattern.Replace(result, "$1");
        result = CodePattern.Replace(result, "$1");
        return result;
    }
}
