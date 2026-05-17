using System.Text;
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

    private enum BlockKind { Empty, Paragraph, Heading, Bullet, Hr, PageBreak, Table }

    private sealed class Block
    {
        public BlockKind Kind;
        public string Text = "";
        public int HeadingLevel;
        public List<List<string>>? TableRows;
        public bool TableHasHeader;
    }

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

        // Phase 1: パース（Word起動前に完了させる）
        var blocks = ParseBlocks(markdown, numberHeadings);
        if (blocks.Count == 0) blocks.Add(new Block { Kind = BlockKind.Empty });

        dynamic? app = null;
        dynamic? doc = null;

        try
        {
            AppLog.Info($"変換開始: output={outputPath} bodyFont={bodyFontName} numberHeadings={numberHeadings} blocks={blocks.Count}");
            progress?.Report(5);

            app = Activator.CreateInstance(wordType)
                ?? throw new InvalidOperationException("Wordを起動できません。");
            app.Visible = false;
            doc = app.Documents.Add();
            progress?.Report(10);

            SetAuthor(doc);

            // Phase 2: 全段落テキストを一括書き込み
            WriteAllText(doc, blocks);
            progress?.Report(40);

            // Phase 3: 段落インデックスでスタイル適用（順方向、テーブルはスキップ）
            ApplyStyles(doc, blocks, bodyFontName, bodyFontSize, progress);
            progress?.Report(70);

            // Phase 4: テーブル挿入（逆順、後続インデックスのシフトを無視できる）
            InsertTables(doc, blocks, bodyFontName, bodyFontSize);
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

    // Phase 1: Markdown を Block リストに変換する（Word に触れない純粋関数）
    private static List<Block> ParseBlocks(string markdown, bool numberHeadings)
    {
        var lines = markdown.Replace("\r\n", "\n").Split('\n');
        var blocks = new List<Block>();
        var headingCounters = new int[3];
        int i = 0;

        while (i < lines.Length)
        {
            var line = lines[i].TrimEnd();

            // テーブルブロック（連続する | 行をまとめる）
            if (IsTableLine(line))
            {
                var tableLines = new List<string>();
                while (i < lines.Length && IsTableLine(lines[i].TrimEnd()))
                {
                    tableLines.Add(lines[i].TrimEnd());
                    i++;
                }
                var hasSeparator = tableLines.Any(IsTableSeparator);
                var dataRows = tableLines
                    .Where(l => !IsTableSeparator(l))
                    .Select(ParseTableRow)
                    .ToList();
                if (dataRows.Count > 0 && dataRows[0].Count > 0)
                {
                    blocks.Add(new Block
                    {
                        Kind = BlockKind.Table,
                        TableRows = dataRows,
                        TableHasHeader = hasSeparator,
                    });
                }
                continue;
            }

            if (string.IsNullOrWhiteSpace(line))
            {
                blocks.Add(new Block { Kind = BlockKind.Empty });
                i++;
                continue;
            }

            var headingMatch = HeadingPattern.Match(line);
            if (headingMatch.Success)
            {
                var level = Math.Min(headingMatch.Groups[1].Value.Length, 3);
                var text = ParseInline(headingMatch.Groups[2].Value);
                if (numberHeadings)
                {
                    headingCounters[level - 1]++;
                    for (int j = level; j < 3; j++) headingCounters[j] = 0;
                    var prefix = level switch
                    {
                        1 => $"{headingCounters[0]}. ",
                        2 => $"{headingCounters[0]}.{headingCounters[1]} ",
                        _ => $"{headingCounters[0]}.{headingCounters[1]}.{headingCounters[2]} ",
                    };
                    text = prefix + text;
                }
                blocks.Add(new Block { Kind = BlockKind.Heading, Text = text, HeadingLevel = level });
                i++;
                continue;
            }

            // 改ページ判定は HR より先（--- 系の重複を避ける）
            if (PageBreakPattern.IsMatch(line))
            {
                blocks.Add(new Block { Kind = BlockKind.PageBreak });
                i++;
                continue;
            }

            var bulletMatch = BulletPattern.Match(line);
            if (bulletMatch.Success)
            {
                blocks.Add(new Block { Kind = BlockKind.Bullet, Text = ParseInline(bulletMatch.Groups[1].Value) });
                i++;
                continue;
            }

            if (HrPattern.IsMatch(line))
            {
                blocks.Add(new Block { Kind = BlockKind.Hr });
                i++;
                continue;
            }

            blocks.Add(new Block { Kind = BlockKind.Paragraph, Text = ParseInline(line) });
            i++;
        }

        return blocks;
    }

    private static string GetParagraphText(Block block) => block.Kind switch
    {
        BlockKind.Heading => block.Text,
        BlockKind.Bullet => block.Text,
        BlockKind.Paragraph => block.Text,
        BlockKind.PageBreak => "\f", // Chr(12) = wdPageBreak
        _ => "", // Empty / Hr / Table はプレースホルダー段落
    };

    // Phase 2: 全段落テキストを 1 回の COM 呼び出しで書き込む
    private static void WriteAllText(dynamic doc, List<Block> blocks)
    {
        var sb = new StringBuilder();
        for (int i = 0; i < blocks.Count; i++)
        {
            sb.Append(GetParagraphText(blocks[i]));
            if (i < blocks.Count - 1) sb.Append('\r'); // 段落マーク
        }
        // doc.Content.Text への代入で既存の空段落を含めて全文を置き換える
        doc.Content.Text = sb.ToString();
    }

    // Phase 3: 既に確定した段落構造に対してスタイル・書式を適用する
    private static void ApplyStyles(
        dynamic doc, List<Block> blocks,
        string bodyFontName, double bodyFontSize,
        IProgress<int>? progress)
    {
        int total = Math.Max(blocks.Count, 1);
        for (int i = 0; i < blocks.Count; i++)
        {
            var block = blocks[i];
            // テーブルは Phase 4 で挿入する
            if (block.Kind == BlockKind.Table) continue;

            dynamic para = doc.Paragraphs[i + 1];

            switch (block.Kind)
            {
                case BlockKind.Heading:
                    // WdBuiltinStyle: Heading N = -(N+1)。フォント上書きせず Word 標準書式を使う
                    para.Range.Style = -(block.HeadingLevel + 1);
                    break;
                case BlockKind.Bullet:
                    para.Range.ListFormat.ApplyBulletDefault();
                    para.Range.Font.Name = bodyFontName;
                    para.Range.Font.Size = bodyFontSize;
                    break;
                case BlockKind.Paragraph:
                    para.Range.Font.Name = bodyFontName;
                    para.Range.Font.Size = bodyFontSize;
                    break;
                case BlockKind.Hr:
                    // -3 = wdBorderBottom, 1 = wdLineStyleSingle
                    para.Borders[-3].LineStyle = 1;
                    break;
                case BlockKind.PageBreak:
                case BlockKind.Empty:
                    // 段落テキストのみで成立（追加書式不要）
                    break;
            }

            progress?.Report(40 + (i + 1) * 30 / total);
        }
    }

    // Phase 4: プレースホルダー段落の位置にテーブルを挿入する
    // 逆順処理により、挿入で発生するインデックスシフトの影響を回避
    private static void InsertTables(dynamic doc, List<Block> blocks, string fontName, double fontSize)
    {
        for (int i = blocks.Count - 1; i >= 0; i--)
        {
            var block = blocks[i];
            if (block.Kind != BlockKind.Table) continue;

            var dataRows = block.TableRows!;
            var colCount = dataRows[0].Count;

            dynamic para = doc.Paragraphs[i + 1];
            dynamic table = doc.Tables.Add(para.Range, dataRows.Count, colCount);
            table.Borders.Enable = 1;

            for (int r = 0; r < dataRows.Count; r++)
            {
                var cells = dataRows[r];
                for (int c = 0; c < Math.Min(cells.Count, colCount); c++)
                {
                    var cell = table.Cell(r + 1, c + 1);
                    cell.Range.Text = cells[c];
                    cell.Range.Font.Name = fontName;
                    cell.Range.Font.Size = fontSize;
                    if (block.TableHasHeader && r == 0)
                        cell.Range.Font.Bold = 1;
                }
            }
        }
    }

    private static void SetAuthor(dynamic doc)
    {
        try { doc.BuiltInDocumentProperties("Author").Value = "md2doc"; }
        catch { }
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
        return parts.Length < 2 ? [] : parts[1..^1].Select(p => ParseInline(p.Trim())).ToList();
    }

    private static void SetHeader(dynamic doc, string headerText, int alignment)
    {
        dynamic header = doc.Sections[1].Headers[1];
        header.Range.Text = headerText;
        header.Range.ParagraphFormat.Alignment = alignment;
    }

    private static void SetFooterPageNumbers(dynamic doc, int alignment)
    {
        dynamic footer = doc.Sections[1].Footers[1];
        dynamic range = footer.Range;
        range.Fields.Add(range, 33); // wdFieldPage

        range = footer.Range;
        range.InsertAfter(" / ");
        range.Collapse(0); // wdCollapseEnd
        range.Fields.Add(range, 26); // wdFieldNumPages

        footer.Range.ParagraphFormat.Alignment = alignment;
    }

    private static string ParseInline(string text)
    {
        // <br> 系を Word のソフトリターン（Chr(11)）に変換
        var result = BrTagPattern.Replace(text, "\v");
        result = BoldPattern.Replace(result, "$1");
        result = ItalicPattern.Replace(result, "$1");
        result = CodePattern.Replace(result, "$1");
        return result;
    }
}
