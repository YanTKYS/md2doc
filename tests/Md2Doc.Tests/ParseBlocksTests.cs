using Xunit;

namespace Md2Doc.Tests;

/// <summary>
/// MarkdownParser.ParseBlocks の単体テスト。
/// Word COM 不要で動作する（純粋関数）。
/// </summary>
public class ParseBlocksTests
{
    // 回帰テスト仕様の必須 Markdown
    private const string RegressionMarkdown = """
        # タイトル
        ## 起動方法
        - アプリを起動する
        - Markdownファイルを選択する
        本文1
        - 箇条書きA
        - 箇条書きB
        本文2
        """;

    [Fact]
    public void ParseBlocks_RegressionMarkdown_AllBlocksPresent()
    {
        var blocks = MarkdownParser.ParseBlocks(RegressionMarkdown, numberHeadings: false);

        // 空行は含まれうるが、実質コンテンツは 8 ブロック必要
        var content = blocks.Where(b => b.Kind != BlockKind.Empty).ToList();

        Assert.Equal(BlockKind.Heading, content[0].Kind);
        Assert.Equal("タイトル", content[0].Text);
        Assert.Equal(1, content[0].HeadingLevel);

        Assert.Equal(BlockKind.Heading, content[1].Kind);
        Assert.Equal("起動方法", content[1].Text);
        Assert.Equal(2, content[1].HeadingLevel);

        Assert.Equal(BlockKind.Bullet, content[2].Kind);
        Assert.Equal("アプリを起動する", content[2].Text);

        Assert.Equal(BlockKind.Bullet, content[3].Kind);
        Assert.Equal("Markdownファイルを選択する", content[3].Text);

        Assert.Equal(BlockKind.Paragraph, content[4].Kind);
        Assert.Equal("本文1", content[4].Text);

        Assert.Equal(BlockKind.Bullet, content[5].Kind);
        Assert.Equal("箇条書きA", content[5].Text);

        Assert.Equal(BlockKind.Bullet, content[6].Kind);
        Assert.Equal("箇条書きB", content[6].Text);

        Assert.Equal(BlockKind.Paragraph, content[7].Kind);
        Assert.Equal("本文2", content[7].Text);
    }

    [Fact]
    public void ParseBlocks_NoBulletsLost_ConsecutiveBullets()
    {
        const string md = "- 1行目\n- 2行目\n- 3行目\n- 4行目";
        var blocks = MarkdownParser.ParseBlocks(md, false);
        var bullets = blocks.Where(b => b.Kind == BlockKind.Bullet).ToList();

        Assert.Equal(4, bullets.Count);
        Assert.Equal("1行目", bullets[0].Text);
        Assert.Equal("2行目", bullets[1].Text);
        Assert.Equal("3行目", bullets[2].Text);
        Assert.Equal("4行目", bullets[3].Text);
    }

    [Fact]
    public void ParseBlocks_HeadingFollowedByBullet_BulletNotLost()
    {
        const string md = "## 見出し\n- 最初の箇条書き\n- 2番目の箇条書き";
        var blocks = MarkdownParser.ParseBlocks(md, false);
        var bullets = blocks.Where(b => b.Kind == BlockKind.Bullet).ToList();

        Assert.Equal(2, bullets.Count);
        Assert.Equal("最初の箇条書き", bullets[0].Text);
        Assert.Equal("2番目の箇条書き", bullets[1].Text);
    }

    [Fact]
    public void ParseBlocks_ParagraphAfterBullet_ParagraphNotLost()
    {
        const string md = "- 箇条書き\n本文段落";
        var blocks = MarkdownParser.ParseBlocks(md, false);
        var para = blocks.FirstOrDefault(b => b.Kind == BlockKind.Paragraph);

        Assert.NotNull(para);
        Assert.Equal("本文段落", para.Text);
    }

    [Fact]
    public void ParseBlocks_NumberHeadings_PrefixAdded()
    {
        const string md = "# 章1\n## 節1\n## 節2\n# 章2\n## 節1";
        var blocks = MarkdownParser.ParseBlocks(md, numberHeadings: true);
        var headings = blocks.Where(b => b.Kind == BlockKind.Heading).ToList();

        Assert.Equal("1. 章1", headings[0].Text);
        Assert.Equal("1.1 節1", headings[1].Text);
        Assert.Equal("1.2 節2", headings[2].Text);
        Assert.Equal("2. 章2", headings[3].Text);
        Assert.Equal("2.1 節1", headings[4].Text);
    }

    [Fact]
    public void ParseBlocks_PageBreak_DetectedCorrectly()
    {
        const string md = "段落1\n<!-- pagebreak -->\n段落2\n---pagebreak---\n段落3";
        var blocks = MarkdownParser.ParseBlocks(md, false);
        var pageBreaks = blocks.Where(b => b.Kind == BlockKind.PageBreak).ToList();

        Assert.Equal(2, pageBreaks.Count);
    }

    [Fact]
    public void ParseBlocks_Table_ParsedCorrectly()
    {
        const string md = "| A | B |\n|---|---|\n| 1 | 2 |";
        var blocks = MarkdownParser.ParseBlocks(md, false);
        var table = blocks.FirstOrDefault(b => b.Kind == BlockKind.Table);

        Assert.NotNull(table);
        Assert.True(table.TableHasHeader);
        Assert.Equal(2, table.TableRows!.Count); // ヘッダー行 + データ行
        Assert.Equal("A", table.TableRows[0][0]);
        Assert.Equal("1", table.TableRows[1][0]);
    }

    [Fact]
    public void ParseBlocks_InlineMarkup_Stripped()
    {
        const string md = "**太字** と *斜体* と `コード`";
        var blocks = MarkdownParser.ParseBlocks(md, false);
        var para = blocks.First(b => b.Kind == BlockKind.Paragraph);

        Assert.Equal("太字 と 斜体 と コード", para.Text);
    }
}
