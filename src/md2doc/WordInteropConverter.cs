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

            doc.SaveAs2(outputPath);
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

        foreach (var rawLine in lines)
        {
            var line = rawLine.TrimEnd();
            if (string.IsNullOrWhiteSpace(line))
            {
                doc.Paragraphs.Add();
                continue;
            }

            if (TryHeading(doc, line) || TryBullet(doc, line))
            {
                continue;
            }

            WriteParagraph(doc, ParseInline(line));
        }
    }

    private static bool TryHeading(dynamic doc, string line)
    {
        var match = HeadingPattern.Match(line);
        if (!match.Success)
        {
            return false;
        }

        var level = match.Groups[1].Value.Length;
        var text = ParseInline(match.Groups[2].Value);
        var para = doc.Paragraphs.Add();
        para.Range.Text = text;
        para.Range.set_Style($"Heading {Math.Min(level, 3)}");
        para.Range.InsertParagraphAfter();
        return true;
    }

    private static bool TryBullet(dynamic doc, string line)
    {
        var match = BulletPattern.Match(line);
        if (!match.Success)
        {
            return false;
        }

        var text = ParseInline(match.Groups[1].Value);
        var para = doc.Paragraphs.Add();
        para.Range.Text = text;
        para.Range.ListFormat.ApplyBulletDefault();
        para.Range.InsertParagraphAfter();
        return true;
    }

    private static void WriteParagraph(dynamic doc, string text)
    {
        var para = doc.Paragraphs.Add();
        para.Range.Text = text;
        para.Range.InsertParagraphAfter();
    }

    private static string ParseInline(string text)
    {
        var result = BoldPattern.Replace(text, "$1");
        result = ItalicPattern.Replace(result, "$1");
        result = CodePattern.Replace(result, "$1");
        return result;
    }
}
