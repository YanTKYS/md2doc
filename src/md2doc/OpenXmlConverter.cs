using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Markdig;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;

using MdBlock = Markdig.Syntax.Block;
using MdTable = Markdig.Extensions.Tables.Table;
using MdTableRow = Markdig.Extensions.Tables.TableRow;
using MdTableCell = Markdig.Extensions.Tables.TableCell;

namespace Md2Doc;

/// <summary>
/// Markdig + Open XML SDK による Markdown → docx 変換 POC 実装。
/// Microsoft Word 不要。WordInteropConverter と共存する比較検証用クラス。
/// </summary>
internal static class OpenXmlConverter
{
    private static readonly Regex PageBreakPattern = new(
        @"^(<!--\s*pagebreak\s*-->|---pagebreak---)$",
        RegexOptions.Compiled | RegexOptions.IgnoreCase);

    private const int BulletNumId = 1;       // numId=1: 箇条書き（全体で共有）
    private const int OrderedAbstractId = 1; // abstractNumId=1: 順序付きリストの雛形

    // ────────────────────────────────────────────────────────────────────
    // Public entry point（WordInteropConverter と同一シグネチャ）
    // ────────────────────────────────────────────────────────────────────

    public static void ConvertToDocx(
        string markdown, string outputPath,
        string bodyFontName, double bodyFontSize,
        bool numberHeadings,
        string? headerText, int headerAlignment,
        bool addPageNumbers, int footerAlignment,
        IProgress<int>? progress = null)
    {
        var pipeline = new MarkdownPipelineBuilder()
            .UseAdvancedExtensions() // Tables, Footnotes, etc.
            .Build();

        var mdDoc = Markdown.Parse(markdown, pipeline);
        progress?.Report(10);

        using var wordDoc = WordprocessingDocument.Create(
            outputPath, WordprocessingDocumentType.Document);

        var mainPart = wordDoc.AddMainDocumentPart();
        var body = new Body();
        mainPart.Document = new Document(body);

        AddStyleDefinitions(mainPart);
        var numberingPart = AddNumberingDefinitions(mainPart);
        progress?.Report(20);

        var ctx = new Ctx(bodyFontName, bodyFontSize, numberHeadings, numberingPart);

        int total = Math.Max(mdDoc.Count, 1);
        int i = 0;
        foreach (var block in mdDoc)
        {
            AppendBlock(body, block, ctx);
            progress?.Report(20 + (++i) * 65 / total);
        }

        // OOXML 仕様: Body の最終要素は Paragraph でなければならない
        if (body.LastChild is not Paragraph)
            body.Append(new Paragraph());

        mainPart.Document.Save();
        progress?.Report(95);
    }

    // ────────────────────────────────────────────────────────────────────
    // 変換コンテキスト（1 変換処理あたりの可変状態）
    // ────────────────────────────────────────────────────────────────────

    private sealed class Ctx
    {
        public string FontName { get; }
        public double FontSize { get; }
        public bool NumberHeadings { get; }
        public int[] HeadingCounters { get; } = new int[3];
        public NumberingDefinitionsPart NumberingPart { get; }
        // numId=1 → 箇条書き固定, numId=2以降 → 順序付きリストごとに割り当て
        public int NextOrderedNumId { get; set; } = 2;

        public Ctx(string fontName, double fontSize, bool numberHeadings,
            NumberingDefinitionsPart numberingPart)
        {
            FontName = fontName;
            FontSize = fontSize;
            NumberHeadings = numberHeadings;
            NumberingPart = numberingPart;
        }
    }

    // ────────────────────────────────────────────────────────────────────
    // スタイル・ナンバリング定義
    // ────────────────────────────────────────────────────────────────────

    private static void AddStyleDefinitions(MainDocumentPart mainPart)
    {
        var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
        var styles = new Styles();

        // Normal（既定段落スタイル）
        var normal = new Style { Type = StyleValues.Paragraph, StyleId = "Normal", Default = true };
        normal.Append(new StyleName { Val = "Normal" });
        styles.Append(normal);

        // 見出し H1〜H3: Word の既定スタイルに準じた太字 + サイズ
        (int level, int halfPt, string name)[] headings =
        [
            (1, 32, "heading 1"),  // 16pt
            (2, 28, "heading 2"),  // 14pt
            (3, 24, "heading 3"),  // 12pt
        ];
        foreach (var (level, halfPt, name) in headings)
        {
            var style = new Style { Type = StyleValues.Paragraph, StyleId = $"Heading{level}" };
            style.Append(new StyleName { Val = name });
            style.Append(new BasedOn { Val = "Normal" });
            var srp = new StyleRunProperties();
            srp.Append(new Bold());
            srp.Append(new FontSize { Val = halfPt.ToString() });
            srp.Append(new FontSizeComplexScript { Val = halfPt.ToString() });
            style.Append(srp);
            styles.Append(style);
        }

        stylesPart.Styles = styles;
    }

    private static NumberingDefinitionsPart AddNumberingDefinitions(MainDocumentPart mainPart)
    {
        var part = mainPart.AddNewPart<NumberingDefinitionsPart>();
        var numbering = new Numbering();

        // abstractNumId=0: 箇条書き（•）
        var absBullet = new AbstractNum { AbstractNumberId = 0 };
        absBullet.Append(new MultiLevelType { Val = MultiLevelValues.SingleLevel });
        var lvlBullet = new Level { LevelIndex = 0 };
        lvlBullet.Append(new StartNumberingValue { Val = 1 });
        lvlBullet.Append(new NumberingFormat { Val = NumberFormatValues.Bullet });
        lvlBullet.Append(new LevelText { Val = "•" });
        lvlBullet.Append(new LevelJustification { Val = LevelJustificationValues.Left });
        absBullet.Append(lvlBullet);
        numbering.Append(absBullet);

        // abstractNumId=1: 順序付きリスト（%1.）
        var absOrdered = new AbstractNum { AbstractNumberId = OrderedAbstractId };
        absOrdered.Append(new MultiLevelType { Val = MultiLevelValues.SingleLevel });
        var lvlOrdered = new Level { LevelIndex = 0 };
        lvlOrdered.Append(new StartNumberingValue { Val = 1 });
        lvlOrdered.Append(new NumberingFormat { Val = NumberFormatValues.Decimal });
        lvlOrdered.Append(new LevelText { Val = "%1." });
        lvlOrdered.Append(new LevelJustification { Val = LevelJustificationValues.Left });
        absOrdered.Append(lvlOrdered);
        numbering.Append(absOrdered);

        // numId=1: 箇条書きインスタンス（文書全体で共有）
        var numBullet = new NumberingInstance { NumberID = BulletNumId };
        numBullet.Append(new AbstractNumId { Val = 0 });
        numbering.Append(numBullet);

        part.Numbering = numbering;
        return part;
    }

    // 順序付きリストが複数ある場合に番号がリセットされるよう、
    // ListBlock ごとに新しい NumberingInstance を生成して追加する。
    private static int AddOrderedListInstance(Ctx ctx)
    {
        int numId = ctx.NextOrderedNumId++;
        var num = new NumberingInstance { NumberID = numId };
        num.Append(new AbstractNumId { Val = OrderedAbstractId });
        // StartOverride で各リストを独立して 1 から開始
        var lvlOverride = new LevelOverride { LevelIndex = 0 };
        lvlOverride.Append(new StartOverrideNumberingValue { Val = 1 });
        num.Append(lvlOverride);
        ctx.NumberingPart.Numbering.Append(num);
        return numId;
    }

    // ────────────────────────────────────────────────────────────────────
    // ブロックの変換ディスパッチ
    // ────────────────────────────────────────────────────────────────────

    private static void AppendBlock(Body body, MdBlock block, Ctx ctx)
    {
        switch (block)
        {
            case HeadingBlock h:
                body.Append(RenderHeading(h, ctx));
                break;

            case ParagraphBlock p:
                var text = ExtractInlines(p.Inline);
                if (PageBreakPattern.IsMatch(text.Trim()))
                    body.Append(RenderPageBreak());
                else
                    body.Append(RenderBodyParagraph(text, ctx));
                break;

            case ListBlock list:
                RenderList(body, list, ctx);
                break;

            case MdTable table:
                body.Append(RenderTable(table, ctx));
                break;

            case HtmlBlock html:
                // <!-- pagebreak --> はブロックコメントとして HtmlBlock に分類される
                if (PageBreakPattern.IsMatch(html.Lines.ToString().Trim()))
                    body.Append(RenderPageBreak());
                break;

            // ThematicBreakBlock (---) は HR として扱い、今回は出力スキップ
        }
    }

    // ────────────────────────────────────────────────────────────────────
    // 各要素のレンダリング
    // ────────────────────────────────────────────────────────────────────

    private static Paragraph RenderHeading(HeadingBlock h, Ctx ctx)
    {
        var level = Math.Min(h.Level, 3);
        var text = ExtractInlines(h.Inline);

        if (ctx.NumberHeadings)
        {
            ctx.HeadingCounters[level - 1]++;
            for (int j = level; j < 3; j++) ctx.HeadingCounters[j] = 0;
            var prefix = level switch
            {
                1 => $"{ctx.HeadingCounters[0]}. ",
                2 => $"{ctx.HeadingCounters[0]}.{ctx.HeadingCounters[1]} ",
                _ => $"{ctx.HeadingCounters[0]}.{ctx.HeadingCounters[1]}.{ctx.HeadingCounters[2]} ",
            };
            text = prefix + text;
        }

        var para = new Paragraph();
        para.Append(new ParagraphProperties(
            new ParagraphStyleId { Val = $"Heading{level}" }));
        // 見出しはフォント名のみ設定（サイズはスタイル定義に従う）
        para.Append(MakeRun(text, ctx.FontName, fontSize: null));
        return para;
    }

    private static Paragraph RenderBodyParagraph(string text, Ctx ctx)
    {
        var para = new Paragraph();
        para.Append(MakeRun(text, ctx.FontName, ctx.FontSize));
        return para;
    }

    private static void RenderList(Body body, ListBlock list, Ctx ctx)
    {
        int numId = list.IsOrdered
            ? AddOrderedListInstance(ctx)
            : BulletNumId;

        foreach (var item in list.OfType<ListItemBlock>())
        {
            var itemText = GetContainerText(item);
            var para = new Paragraph();
            para.Append(new ParagraphProperties(
                new NumberingProperties(
                    new NumberingLevelReference { Val = 0 },
                    new NumberingId { Val = numId })));
            para.Append(MakeRun(itemText, ctx.FontName, ctx.FontSize));
            body.Append(para);
        }
    }

    private static Table RenderTable(MdTable mdTable, Ctx ctx)
    {
        var table = new Table();
        table.Append(new TableProperties(
            new TableBorders(
                new TopBorder { Val = BorderValues.Single, Size = 4 },
                new BottomBorder { Val = BorderValues.Single, Size = 4 },
                new LeftBorder { Val = BorderValues.Single, Size = 4 },
                new RightBorder { Val = BorderValues.Single, Size = 4 },
                new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
                new InsideVerticalBorder { Val = BorderValues.Single, Size = 4 })));

        foreach (var row in mdTable.OfType<MdTableRow>())
        {
            var tr = new TableRow();
            foreach (var cell in row.OfType<MdTableCell>())
            {
                var tc = new TableCell();
                var para = new Paragraph();
                var run = MakeRun(GetContainerText(cell), ctx.FontName, ctx.FontSize);
                if (row.IsHeader)
                    run.GetFirstChild<RunProperties>()!.Append(new Bold());
                para.Append(run);
                tc.Append(para);
                tr.Append(tc);
            }
            table.Append(tr);
        }

        return table;
    }

    private static Paragraph RenderPageBreak()
    {
        var para = new Paragraph();
        para.Append(new Run(new Break { Type = BreakValues.Page }));
        return para;
    }

    // ────────────────────────────────────────────────────────────────────
    // Run / テキスト生成ヘルパー
    // ────────────────────────────────────────────────────────────────────

    // fontSize が null のとき（見出し用）はフォントサイズをスタイルに委ねる。
    private static Run MakeRun(string text, string fontName, double? fontSize)
    {
        var run = new Run();
        var rp = new RunProperties();
        rp.Append(new RunFonts
        {
            Ascii = fontName,
            HighAnsi = fontName,
            EastAsia = fontName,
            ComplexScript = fontName,
        });
        if (fontSize is double sz)
        {
            var halfPt = ((int)(sz * 2)).ToString();
            rp.Append(new FontSize { Val = halfPt });
            rp.Append(new FontSizeComplexScript { Val = halfPt });
        }
        run.Append(rp);

        // \n を Break 要素に変換（<br> タグ等に相当）
        var lines = text.Split('\n');
        for (int i = 0; i < lines.Length; i++)
        {
            if (i > 0) run.Append(new Break());
            if (lines[i].Length > 0)
                run.Append(new Text(lines[i]) { Space = SpaceProcessingModeValues.Preserve });
        }
        return run;
    }

    // Markdig の ContainerInline からプレーンテキストを再帰的に抽出する。
    // インライン書式（太字・斜体・コード）はマーカー除去してテキストのみ取得。
    private static string ExtractInlines(ContainerInline? container)
    {
        if (container is null) return "";
        var sb = new StringBuilder();
        foreach (var inline in container)
        {
            switch (inline)
            {
                case LiteralInline lit:
                    sb.Append(lit.Content.ToString());
                    break;
                case CodeInline code:
                    sb.Append(code.Content);
                    break;
                case LineBreakInline:
                    sb.Append('\n');
                    break;
                case ContainerInline ci: // EmphasisInline 等
                    sb.Append(ExtractInlines(ci));
                    break;
            }
        }
        return sb.ToString();
    }

    // ListItemBlock・TableCell 等の ContainerBlock からテキストを再帰抽出する。
    private static string GetContainerText(ContainerBlock container)
    {
        var sb = new StringBuilder();
        foreach (var child in container)
        {
            if (child is LeafBlock leaf && leaf.Inline is not null)
                sb.Append(ExtractInlines(leaf.Inline));
            else if (child is ContainerBlock nested)
                sb.Append(GetContainerText(nested));
        }
        return sb.ToString();
    }
}
