using System.Text.RegularExpressions;

namespace Md2Doc;

internal enum BlockKind { Empty, Paragraph, Heading, Bullet, Hr, PageBreak, Table }

internal sealed class Block
{
    public BlockKind Kind;
    public string Text = "";
    public int HeadingLevel;
    public List<List<string>>? TableRows;
    public bool TableHasHeader;
}

internal static class MarkdownParser
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

    internal static List<Block> ParseBlocks(string markdown, bool numberHeadings)
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

    internal static string ParseInline(string text)
    {
        // <br> 系を Word のソフトリターン（Chr(11)）に変換
        var result = BrTagPattern.Replace(text, "\v");
        result = BoldPattern.Replace(result, "$1");
        result = ItalicPattern.Replace(result, "$1");
        result = CodePattern.Replace(result, "$1");
        return result;
    }

    internal static bool IsTableLine(string line)
    {
        var t = line.Trim();
        return t.Length > 2 && t.StartsWith("|") && t.EndsWith("|");
    }

    internal static bool IsTableSeparator(string line) =>
        line.All(c => c is '-' or ':' or '|' or ' ');

    internal static List<string> ParseTableRow(string line)
    {
        var parts = line.Split('|');
        return parts.Length < 2 ? [] : parts[1..^1].Select(p => ParseInline(p.Trim())).ToList();
    }
}
