using System.Text;

namespace Md2Doc;

internal static class WordInteropConverter
{
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
        var blocks = MarkdownParser.ParseBlocks(markdown, numberHeadings);
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

        // パス 1: 各ブロックの段落スタイル・フォントを適用
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
                    // -47 = wdStyleListBullet。ApplyBulletDefault() と異なりトグルではなく確定的な設定
                    para.Range.Style = -47;
                    para.Range.Font.Name = bodyFontName;
                    para.Range.Font.Size = bodyFontSize;
                    break;
                case BlockKind.Paragraph:
                    para.Range.Style = -1; // wdStyleNormal
                    para.Range.Font.Name = bodyFontName;
                    para.Range.Font.Size = bodyFontSize;
                    break;
                case BlockKind.Empty:
                    para.Range.Style = -1;
                    break;
                case BlockKind.Hr:
                    para.Range.Style = -1;
                    // -3 = wdBorderBottom, 1 = wdLineStyleSingle
                    para.Borders[-3].LineStyle = 1;
                    break;
                case BlockKind.PageBreak:
                    // \f で段落テキストが設定済み、追加書式不要
                    break;
            }

            progress?.Report(40 + (i + 1) * 20 / total);
        }

        // パス 2: 隣接箇条書きから自動伝播したリスト書式を非箇条書き段落から除去
        // Word は wdStyleListBullet 適用時に隣接段落へリスト書式を直接適用形式で伝播することがあり、
        // Style 設定だけでは消えない。パス 1 完了後に一括除去することで両方向の伝播を確実に拾う。
        for (int i = 0; i < blocks.Count; i++)
        {
            var block = blocks[i];
            if (block.Kind is BlockKind.Bullet or BlockKind.Table or BlockKind.PageBreak)
                continue;

            try
            {
                dynamic para = doc.Paragraphs[i + 1];
                // 0 = wdListNoNumbering。既にリスト書式が無ければ何もしない
                if ((int)para.Range.ListFormat.ListType != 0)
                    para.Range.ListFormat.RemoveNumbers();
            }
            catch (Exception ex)
            {
                AppLog.Error($"リスト書式の除去失敗 block[{i}]={block.Kind}", ex);
            }

            progress?.Report(60 + (i + 1) * 10 / total);
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
}
