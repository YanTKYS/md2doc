using System.IO.Compression;
using System.Xml.Linq;

namespace Md2Doc.Tests;

/// <summary>
/// .docx（ZIP）から word/document.xml を取り出してテキストを検査するヘルパー。
/// Word COM 不要で動作する。
/// </summary>
internal static class DocxInspector
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    /// <summary>
    /// 全段落のテキストを順番に返す。
    /// 各段落内の w:t 要素を結合し、1 要素 = 1 段落の文字列とする。
    /// </summary>
    internal static IReadOnlyList<string> ExtractParagraphTexts(string docxPath)
    {
        using var zip = ZipFile.OpenRead(docxPath);
        var entry = zip.GetEntry("word/document.xml")
            ?? throw new InvalidOperationException($"{docxPath} に word/document.xml が見つかりません。");

        using var stream = entry.Open();
        var xml = XDocument.Load(stream);

        return xml.Descendants(W + "p")
            .Select(p => string.Concat(p.Descendants(W + "t").Select(t => t.Value)))
            .Where(t => t.Length > 0)
            .ToList();
    }

    /// <summary>
    /// 指定テキストが expectedTexts の順序通りに段落として存在するか検証する。
    /// 検証失敗時は最初に見つからなかったテキストとその前後の段落情報を返す。
    /// </summary>
    internal static (bool ok, string? failMessage) VerifyOrder(
        IReadOnlyList<string> paragraphTexts, IReadOnlyList<string> expectedTexts)
    {
        int lastIndex = -1;
        foreach (var expected in expectedTexts)
        {
            int found = -1;
            for (int i = lastIndex + 1; i < paragraphTexts.Count; i++)
            {
                if (paragraphTexts[i].Contains(expected))
                {
                    found = i;
                    break;
                }
            }
            if (found < 0)
            {
                var context = string.Join(", ",
                    paragraphTexts.Select((t, idx) => $"[{idx}]\"{t}\"").Take(20));
                return (false,
                    $"テキスト '{expected}' が document.xml に存在しないか、" +
                    $"'{(lastIndex >= 0 ? expectedTexts[Array.IndexOf(expectedTexts.ToArray(), expected) - 1] : "(先頭)")}' より後に見つかりません。" +
                    $"\n実際の段落一覧: {context}");
            }
            lastIndex = found;
        }
        return (true, null);
    }
}
