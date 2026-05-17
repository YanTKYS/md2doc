using System.IO.Compression;
using System.Xml.Linq;

namespace Md2Doc.Tests;

/// <summary>
/// .docx（ZIP）から XML パーツを取り出して検査するヘルパー。
/// Word COM 不要で動作する。
/// </summary>
internal static class DocxInspector
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    // ──────────────────────────────────────────────────────────────────────
    // 段落テキスト抽出
    // ──────────────────────────────────────────────────────────────────────

    /// <summary>
    /// 本文の全段落テキストを順番に返す。
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
    /// ヘッダーパーツ（word/header*.xml）の全段落テキストを返す。
    /// </summary>
    internal static IReadOnlyList<string> ExtractHeaderTexts(string docxPath)
        => ExtractPartsTexts(docxPath, "word/header");

    // ──────────────────────────────────────────────────────────────────────
    // 順序検証
    // ──────────────────────────────────────────────────────────────────────

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

    // ──────────────────────────────────────────────────────────────────────
    // 書式・要素検査（v0.5.2 追加）
    // ──────────────────────────────────────────────────────────────────────

    /// <summary>
    /// 指定テキストを含む w:r の RunProperties に指定プロパティ要素が存在するか確認する。
    /// 例: propertyLocalName="b" → 太字, "i" → 斜体
    /// </summary>
    internal static bool HasRunWithTextAndProperty(string docxPath, string text, string propertyLocalName)
    {
        var xdoc = LoadDocumentXml(docxPath);
        return xdoc.Descendants(W + "r").Any(run =>
        {
            var runText = string.Concat(run.Descendants(W + "t").Select(t => t.Value));
            return runText.Contains(text)
                && run.Elements(W + "rPr").Elements(W + propertyLocalName).Any();
        });
    }

    /// <summary>
    /// 指定テキストを含む w:r の RunFonts の Ascii / HighAnsi に指定フォントが設定されているか確認する。
    /// インラインコードの等幅フォント検証に使用する。
    /// </summary>
    internal static bool HasRunWithFont(string docxPath, string text, string fontName)
    {
        var xdoc = LoadDocumentXml(docxPath);
        return xdoc.Descendants(W + "r").Any(run =>
        {
            var runText = string.Concat(run.Descendants(W + "t").Select(t => t.Value));
            if (!runText.Contains(text)) return false;
            var rFonts = run.Elements(W + "rPr").Elements(W + "rFonts").FirstOrDefault();
            return rFonts is not null &&
                ((string?)rFonts.Attribute(W + "ascii") == fontName ||
                 (string?)rFonts.Attribute(W + "hAnsi") == fontName);
        });
    }

    /// <summary>
    /// フッターパーツに PAGE フィールドが存在するか確認する。
    /// </summary>
    internal static bool FooterHasPageNumber(string docxPath)
    {
        using var zip = ZipFile.OpenRead(docxPath);
        return zip.Entries
            .Where(e => e.FullName.StartsWith("word/footer", StringComparison.OrdinalIgnoreCase)
                     && e.FullName.EndsWith(".xml"))
            .Any(entry =>
            {
                using var stream = entry.Open();
                var xdoc = XDocument.Load(stream);
                return xdoc.Descendants(W + "instrText")
                    .Any(t => t.Value.Contains("PAGE", StringComparison.OrdinalIgnoreCase));
            });
    }

    /// <summary>
    /// 本文に水平線（段落の下罫線）が存在するか確認する。
    /// </summary>
    internal static bool DocumentHasHorizontalRule(string docxPath)
    {
        var xdoc = LoadDocumentXml(docxPath);
        return xdoc.Descendants(W + "pBdr")
            .Any(pBdr => pBdr.Element(W + "bottom") is not null);
    }

    /// <summary>
    /// 本文にソフトリターン（テキスト行内改行の w:br）が存在するか確認する。
    /// ページ改ページ（w:type="page"）は除外する。
    /// </summary>
    internal static bool DocumentHasSoftReturn(string docxPath)
    {
        var xdoc = LoadDocumentXml(docxPath);
        return xdoc.Descendants(W + "br")
            .Any(br => (string?)br.Attribute(W + "type") != "page");
    }

    // ──────────────────────────────────────────────────────────────────────
    // 内部ヘルパー
    // ──────────────────────────────────────────────────────────────────────

    private static XDocument LoadDocumentXml(string docxPath)
    {
        using var zip = ZipFile.OpenRead(docxPath);
        var entry = zip.GetEntry("word/document.xml")
            ?? throw new InvalidOperationException($"{docxPath} に word/document.xml が見つかりません。");
        using var stream = entry.Open();
        return XDocument.Load(stream);
    }

    private static IReadOnlyList<string> ExtractPartsTexts(string docxPath, string namePrefix)
    {
        using var zip = ZipFile.OpenRead(docxPath);
        var texts = new List<string>();
        foreach (var entry in zip.Entries.Where(e =>
            e.FullName.StartsWith(namePrefix, StringComparison.OrdinalIgnoreCase)
            && e.FullName.EndsWith(".xml")))
        {
            using var stream = entry.Open();
            var xdoc = XDocument.Load(stream);
            texts.AddRange(
                xdoc.Descendants(W + "p")
                    .Select(p => string.Concat(p.Descendants(W + "t").Select(t => t.Value)))
                    .Where(t => t.Length > 0));
        }
        return texts;
    }
}
