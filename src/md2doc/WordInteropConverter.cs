using System.Text.RegularExpressions;

namespace Md2Doc;

internal static class WordInteropConverter
{
    private static readonly Regex HeadingPattern = new(@"^(#{1,6})\s+(.+)$", RegexOptions.Compiled);
    private static readonly Regex BulletPattern = new(@"^[-*]\s+(.+)$", RegexOptions.Compiled);
    private static readonly Regex BoldPattern = new(@"\*\*(.+?)\*\*", RegexOptions.Compiled);
    private static readonly Regex ItalicPattern = new(@"\*(.+?)\*", RegexOptions.Compiled);
    private static readonly Regex CodePattern = new(@"`(.+?)`", RegexOptions.Compiled);

    public static void ConvertToDocx(string markdown, string outputPath, string fontName, double fontSize)
    {
        var wordType = Type.GetTypeFromProgID("Word.Application")
            ?? throw new InvalidOperationException("Microsoft Word が利用できません。");

        dynamic? app = null;
        dynamic? doc = null;

        try
        {
            app = Activator.CreateInstance(wordType)
                ?? throw new InvalidOperationException("Wordを起動できません。");
            app.Visible = false;
            doc = app.Documents.Add();

            WriteMarkdown(doc, markdown);

            doc.Content.Font.Name = fontName;
            doc.Content.Font.Size = fontSize;

            doc.SaveAs2(outputPath, 12); // 12 = wdFormatXMLDocument (.docx)
        }
        finally
        {
            if (doc is not null)
            {
                doc.Close(false);
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(doc);
            }

            if (app is not null)
            {
                app.Quit(false);
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(app);
            }
        }
    }

    private static void WriteMarkdown(dynamic doc, string markdown)
    {
        var lines = markdown.Replace("\r\n", "\n").Split('\n');
        bool firstParagraphUsed = false;

        foreach (var rawLine in lines)
        {
            var line = rawLine.TrimEnd();
            if (string.IsNullOrWhiteSpace(line))
            {
                if (firstParagraphUsed)
                {
                    doc.Paragraphs.Add();
                }
                continue;
            }

            // 新規ドキュメントの最初の空段落を再利用することで先頭の余分な空行を防ぐ
            dynamic para = firstParagraphUsed ? doc.Paragraphs.Add() : doc.Paragraphs[1];
            firstParagraphUsed = true;

            if (!TryHeading(para, line) && !TryBullet(para, line))
            {
                WriteParagraph(para, ParseInline(line));
            }
        }
    }

    private static bool TryHeading(dynamic para, string line)
    {
        var match = HeadingPattern.Match(line);
        if (!match.Success)
        {
            return false;
        }

        var level = match.Groups[1].Value.Length;
        para.Range.Text = ParseInline(match.Groups[2].Value);
        para.Range.set_Style($"Heading {Math.Min(level, 3)}");
        return true;
    }

    private static bool TryBullet(dynamic para, string line)
    {
        var match = BulletPattern.Match(line);
        if (!match.Success)
        {
            return false;
        }

        para.Range.Text = ParseInline(match.Groups[1].Value);
        para.Range.ListFormat.ApplyBulletDefault();
        return true;
    }

    private static void WriteParagraph(dynamic para, string text)
    {
        para.Range.Text = text;
    }

    private static string ParseInline(string text)
    {
        var result = BoldPattern.Replace(text, "$1");
        result = ItalicPattern.Replace(result, "$1");
        result = CodePattern.Replace(result, "$1");
        return result;
    }
}
