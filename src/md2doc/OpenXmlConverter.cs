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
/// Markdig + Open XML SDK による Markdown → docx 変換クラス（v0.5.2）。
/// Microsoft Word 不要。WordInteropConverter と共存する本実装候補クラス。
/// </summary>
internal static class OpenXmlConverter
{
    private static readonly Regex PageBreakPattern = new(
        @"^(<!--\s*pagebreak\s*-->|---pagebreak---)$",
        RegexOptions.Compiled | RegexOptions.IgnoreCase);

    private const int BulletNumId = 1;       // numId=1: 箇条書き（文書全体で共有）
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
        var pipeline = new MarkdownPipelineBuilder().UseAdvancedExtensions().Build();
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

        // OOXML: Body の最終要素は SectionProperties の直前に Paragraph が必要
        if (body.LastChild is not Paragraph)
            body.Append(new Paragraph());

        bool hasSect = !string.IsNullOrEmpty(headerText) || addPageNumbers;
        if (hasSect)
        {
            var sectPr = new SectionProperties();
            if (!string.IsNullOrEmpty(headerText))
            {
                var hp = mainPart.AddNewPart<HeaderPart>();
                hp.Header = BuildHeader(headerText, headerAlignment, bodyFontName, bodyFontSize);
                hp.Header.Save();
                sectPr.Append(new HeaderReference
                {
                    Type = HeaderFooterValues.Default,
                    Id = mainPart.GetIdOfPart(hp),
                });
            }
            if (addPageNumbers)
            {
                var fp = mainPart.AddNewPart<FooterPart>();
                fp.Footer = BuildFooter(footerAlignment, bodyFontName, bodyFontSize);
                fp.Footer.Save();
                sectPr.Append(new FooterReference
                {
                    Type = HeaderFooterValues.Default,
                    Id = mainPart.GetIdOfPart(fp),
                });
            }
            body.Append(sectPr);
        }

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

        var normal = new Style { Type = StyleValues.Paragraph, StyleId = "Normal", Default = true };
        normal.Append(new StyleName { Val = "Normal" });
        styles.Append(normal);

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
        var lvlOverride = new LevelOverride { LevelIndex = 0 };
        lvlOverride.Append(new StartOverrideNumberingValue { Val = 1 });
        num.Append(lvlOverride);
        ctx.NumberingPart.Numbering.Append(num);
        return numId;
    }

    // ────────────────────────────────────────────────────────────────────
    // ヘッダー・フッタービルダー
    // ────────────────────────────────────────────────────────────────────

    private static Header BuildHeader(string text, int alignment, string fontName, double fontSize)
    {
        var header = new Header();
        var para = new Paragraph();
        para.Append(new ParagraphProperties(new Justification { Val = ToJustification(alignment) }));
        para.Append(MakePlainRun(text, fontName, fontSize));
        header.Append(para);
        return header;
    }

    private static Footer BuildFooter(int alignment, string fontName, double fontSize)
    {
        var footer = new Footer();
        var para = new Paragraph();
        para.Append(new ParagraphProperties(new Justification { Val = ToJustification(alignment) }));

        // ページ番号フィールド: BEGIN — instrText " PAGE " — END
        var rp = MakeRunProperties(fontName, fontSize, bold: false, italic: false, code: false);
        var runBegin = new Run();
        runBegin.Append((RunProperties)rp.CloneNode(true));
        runBegin.Append(new FieldChar { FieldCharType = FieldCharValues.Begin });

        var runInstr = new Run();
        runInstr.Append((RunProperties)rp.CloneNode(true));
        var instrText = new FieldCode(" PAGE ") { Space = SpaceProcessingModeValues.Preserve };
        runInstr.Append(instrText);

        var runEnd = new Run();
        runEnd.Append((RunProperties)rp.CloneNode(true));
        runEnd.Append(new FieldChar { FieldCharType = FieldCharValues.End });

        para.Append(runBegin, runInstr, runEnd);
        footer.Append(para);
        return footer;
    }

    private static JustificationValues ToJustification(int alignment) => alignment switch
    {
        1 => JustificationValues.Center,
        2 => JustificationValues.Right,
        _ => JustificationValues.Left,
    };

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
                var text = ExtractText(p.Inline);
                if (PageBreakPattern.IsMatch(text.Trim()))
                    body.Append(RenderPageBreak());
                else
                    body.Append(RenderBodyParagraph(p.Inline, ctx));
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

            case ThematicBreakBlock:
                body.Append(RenderHorizontalRule());
                break;
        }
    }

    // ────────────────────────────────────────────────────────────────────
    // 各要素のレンダリング
    // ────────────────────────────────────────────────────────────────────

    private static Paragraph RenderHeading(HeadingBlock h, Ctx ctx)
    {
        var styleLevel = Math.Min(h.Level, 3);
        var para = new Paragraph();
        para.Append(new ParagraphProperties(new ParagraphStyleId { Val = $"Heading{styleLevel}" }));

        // 見出し番号は H1〜H3 のみ付与。H4 以上は番号カウンタを破壊しないよう
        // 番号付与とカウンタ更新の両方をスキップする（スタイルは H3 として描画される）。
        if (ctx.NumberHeadings && h.Level <= 3)
        {
            ctx.HeadingCounters[styleLevel - 1]++;
            for (int j = styleLevel; j < 3; j++) ctx.HeadingCounters[j] = 0;
            var prefix = styleLevel switch
            {
                1 => $"{ctx.HeadingCounters[0]}. ",
                2 => $"{ctx.HeadingCounters[0]}.{ctx.HeadingCounters[1]} ",
                _ => $"{ctx.HeadingCounters[0]}.{ctx.HeadingCounters[1]}.{ctx.HeadingCounters[2]} ",
            };
            para.Append(MakePlainRun(prefix, ctx.FontName, fontSize: null));
        }

        // 見出し本文のインライン書式を反映（太字・斜体等）
        AppendInlineRuns(para, h.Inline, ctx.FontName, fontSize: null);
        return para;
    }

    // インライン書式付き本文段落（主要パス）
    private static Paragraph RenderBodyParagraph(ContainerInline? inline, Ctx ctx)
    {
        var para = new Paragraph();
        AppendInlineRuns(para, inline, ctx.FontName, ctx.FontSize);
        return para;
    }

    // プレーンテキスト本文段落（リスト lazy-continuation パス）
    private static Paragraph RenderBodyParagraph(string text, Ctx ctx)
    {
        var para = new Paragraph();
        para.Append(MakePlainRun(text, ctx.FontName, ctx.FontSize));
        return para;
    }

    private static void RenderList(Body body, ListBlock list, Ctx ctx)
    {
        int numId = list.IsOrdered ? AddOrderedListInstance(ctx) : BulletNumId;

        foreach (var item in list.OfType<ListItemBlock>())
        {
            bool first = true;
            foreach (var child in item)
            {
                var childText = child is LeafBlock leaf && leaf.Inline is not null
                    ? ExtractText(leaf.Inline)
                    : child is ContainerBlock nested ? GetContainerText(nested) : "";

                // Markdig は空行なしの後続行をリスト項目内へ lazy continuation として
                // 取り込む。改行で分割し、最初の行のみをリスト項目テキストとし、
                // 後続行は独立した本文段落として出力する。
                var lines = childText.Split('\n');
                foreach (var line in lines)
                {
                    if (first)
                    {
                        var para = new Paragraph();
                        para.Append(new ParagraphProperties(
                            new NumberingProperties(
                                new NumberingLevelReference { Val = 0 },
                                new NumberingId { Val = numId })));
                        para.Append(MakePlainRun(line, ctx.FontName, ctx.FontSize));
                        body.Append(para);
                        first = false;
                    }
                    else
                    {
                        body.Append(RenderBodyParagraph(line, ctx));
                    }
                }
            }
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
                var cellLeaf = cell.OfType<LeafBlock>().FirstOrDefault();
                if (cellLeaf?.Inline is not null)
                    AppendInlineRuns(para, cellLeaf.Inline, ctx.FontName, ctx.FontSize, bold: row.IsHeader);
                else
                    para.Append(MakePlainRun(GetContainerText(cell), ctx.FontName, ctx.FontSize, bold: row.IsHeader));
                tc.Append(para);
                tr.Append(tc);
            }
            table.Append(tr);
        }
        return table;
    }

    private static Paragraph RenderHorizontalRule()
    {
        var para = new Paragraph();
        var pProps = new ParagraphProperties();
        var borders = new ParagraphBorders();
        borders.Append(new BottomBorder { Val = BorderValues.Single, Size = 6, Space = 1, Color = "000000" });
        pProps.Append(borders);
        para.Append(pProps);
        // Run なし段落は一部のリーダー（LibreOffice 等）で罫線が描画されないため空 Run を追加する
        para.Append(new Run());
        return para;
    }

    private static Paragraph RenderPageBreak()
    {
        var para = new Paragraph();
        para.Append(new Run(new Break { Type = BreakValues.Page }));
        return para;
    }

    // ────────────────────────────────────────────────────────────────────
    // インラインレンダリング（インライン書式対応）
    // ────────────────────────────────────────────────────────────────────

    // Markdig の ContainerInline を走査し、書式を保持した Run 要素を段落に追加する。
    // bold / italic は親 EmphasisInline から継承される。
    private static void AppendInlineRuns(
        Paragraph para, ContainerInline? container,
        string fontName, double? fontSize,
        bool bold = false, bool italic = false)
    {
        if (container is null) return;
        foreach (var inline in container)
        {
            switch (inline)
            {
                case LiteralInline lit:
                    para.Append(MakeFormattedRun(lit.Content.ToString(), fontName, fontSize, bold, italic, code: false));
                    break;

                case CodeInline code:
                    // インラインコードは等幅フォント + bold/italic をリセット
                    para.Append(MakeFormattedRun(code.Content, fontName, fontSize, bold: false, italic: false, code: true));
                    break;

                case LineBreakInline:
                    // 硬改行（`  \n` や HardlineBreaks 設定）→ ソフトリターン
                    para.Append(new Run(new Break()));
                    break;

                case HtmlInline html when html.Tag.StartsWith("<br", StringComparison.OrdinalIgnoreCase):
                    // <br> / <br/> タグ → ソフトリターン
                    para.Append(new Run(new Break()));
                    break;

                case EmphasisInline em:
                    // DelimiterCount >= 2 → 太字, == 1 → 斜体（** / __ / * / _）
                    AppendInlineRuns(para, em, fontName, fontSize,
                        bold: bold || em.DelimiterCount >= 2,
                        italic: italic || em.DelimiterCount == 1);
                    break;

                case ContainerInline ci:
                    AppendInlineRuns(para, ci, fontName, fontSize, bold, italic);
                    break;
            }
        }
    }

    // ────────────────────────────────────────────────────────────────────
    // Run / RunProperties ヘルパー
    // ────────────────────────────────────────────────────────────────────

    // インライン書式付き Run（body 段落・見出し・表セルで使用）
    private static Run MakeFormattedRun(
        string text, string fontName, double? fontSize,
        bool bold, bool italic, bool code)
    {
        var run = new Run();
        run.Append(MakeRunProperties(fontName, fontSize, bold, italic, code));
        if (text.Length > 0)
            run.Append(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
        return run;
    }

    // プレーンテキスト Run（リスト項目・継続段落・ヘッダーで使用）
    private static Run MakePlainRun(string text, string fontName, double? fontSize, bool bold = false)
    {
        var run = new Run();
        run.Append(MakeRunProperties(fontName, fontSize, bold, italic: false, code: false));
        if (text.Length > 0)
            run.Append(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
        return run;
    }

    private static RunProperties MakeRunProperties(
        string fontName, double? fontSize,
        bool bold, bool italic, bool code)
    {
        string actualFont = code ? "Courier New" : fontName;
        var rp = new RunProperties();
        rp.Append(new RunFonts
        {
            Ascii = actualFont,
            HighAnsi = actualFont,
            EastAsia = fontName,   // 東アジア文字は常に指定フォントを適用
            ComplexScript = actualFont,
        });
        if (fontSize is double sz)
        {
            var halfPt = ((int)(sz * 2)).ToString();
            rp.Append(new FontSize { Val = halfPt });
            rp.Append(new FontSizeComplexScript { Val = halfPt });
        }
        if (bold) rp.Append(new Bold());
        if (italic) rp.Append(new Italic());
        return rp;
    }

    // ────────────────────────────────────────────────────────────────────
    // テキスト抽出（改ページ判定・リスト lazy-continuation 分割専用）
    // ────────────────────────────────────────────────────────────────────

    // プレーンテキストのみを再帰抽出する。インライン書式情報は失われる。
    // ページブレークパターン検出とリスト項目の行分割にのみ使用する。
    private static string ExtractText(ContainerInline? container)
    {
        if (container is null) return "";
        var sb = new StringBuilder();
        foreach (var inline in container)
        {
            switch (inline)
            {
                case LiteralInline lit: sb.Append(lit.Content.ToString()); break;
                case CodeInline code: sb.Append(code.Content); break;
                case LineBreakInline: sb.Append('\n'); break;
                case ContainerInline ci: sb.Append(ExtractText(ci)); break;
            }
        }
        return sb.ToString();
    }

    // ListItemBlock・TableCell 等の ContainerBlock からプレーンテキストを再帰抽出する。
    private static string GetContainerText(ContainerBlock container)
    {
        var sb = new StringBuilder();
        foreach (var child in container)
        {
            if (child is LeafBlock leaf && leaf.Inline is not null)
                sb.Append(ExtractText(leaf.Inline));
            else if (child is ContainerBlock nested)
                sb.Append(GetContainerText(nested));
        }
        return sb.ToString();
    }
}
