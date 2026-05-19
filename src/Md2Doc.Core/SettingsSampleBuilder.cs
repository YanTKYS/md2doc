using System.Text;

namespace Md2Doc.Core;

/// <summary>
/// 現在の設定を人間が読みやすいサンプルテキストに変換する。
/// Word 出力との完全一致は保証しない（設定確認用）。
/// </summary>
public static class SettingsSampleBuilder
{
    /// <param name="headerMode">0=なし 1=ファイル名 2=自由記入</param>
    /// <param name="headerCustomText">自由記入テキスト（headerMode==2 のとき使用）</param>
    /// <param name="inputFileName">ファイル名（headerMode==1 のとき使用、null なら未選択を示す）</param>
    /// <param name="headerAlignment">0=左 1=中央 2=右</param>
    /// <param name="numberHeadings">見出し番号を付与するか</param>
    /// <param name="footerPageNumber">ページ番号を挿入するか</param>
    public static string Build(
        int headerMode,
        string headerCustomText,
        string? inputFileName,
        int headerAlignment,
        bool numberHeadings,
        bool footerPageNumber)
    {
        var sb = new StringBuilder();

        sb.AppendLine("【ヘッダー】");
        if (headerMode == 0)
        {
            sb.AppendLine("（なし）");
        }
        else
        {
            var text = headerMode == 1
                ? (string.IsNullOrWhiteSpace(inputFileName) ? "（ファイル名）" : inputFileName)
                : (string.IsNullOrWhiteSpace(headerCustomText) ? "（未入力）" : headerCustomText);

            sb.AppendLine($"左：{(headerAlignment == 0 ? text : "（なし）")}");
            sb.AppendLine($"中央：{(headerAlignment == 1 ? text : "（なし）")}");
            sb.AppendLine($"右：{(headerAlignment == 2 ? text : "（なし）")}");
        }

        sb.AppendLine();

        sb.AppendLine("【見出し番号】");
        if (!numberHeadings)
        {
            sb.AppendLine("（なし）");
        }
        else
        {
            sb.AppendLine("1. 大見出し");
            sb.AppendLine("1.1 中見出し");
            sb.AppendLine("1.1.1 小見出し");
        }

        sb.AppendLine();

        sb.AppendLine("【ページ番号】");
        sb.Append(footerPageNumber ? "1 / 5" : "（なし）");

        return sb.ToString();
    }
}
