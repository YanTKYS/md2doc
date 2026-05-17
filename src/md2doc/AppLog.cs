using System.Text;

namespace Md2Doc;

internal static class AppLog
{
    private static readonly string FilePath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "md2doc", "app.log");

    private const long MaxBytes = 1024 * 512; // 512 KB

    public static void Info(string message) => Write("INFO", message, null);
    public static void Error(string message, Exception ex) => Write("ERROR", message, ex);

    private static void Write(string level, string message, Exception? ex)
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(FilePath)!);

            var sb = new StringBuilder();
            sb.Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            sb.Append($" [{level}] ");
            sb.AppendLine(message);
            if (ex is not null)
            {
                sb.AppendLine(ex.ToString());
            }

            File.AppendAllText(FilePath, sb.ToString(), Encoding.UTF8);
            TrimIfNeeded();
        }
        catch
        {
            // ログ失敗は変換結果に影響しない
        }
    }

    private static void TrimIfNeeded()
    {
        try
        {
            var info = new FileInfo(FilePath);
            if (info.Length <= MaxBytes) return;

            // ファイルが上限を超えたら後半だけ残す
            var text = File.ReadAllText(FilePath, Encoding.UTF8);
            var half = text.Length / 2;
            var cutAt = text.IndexOf('\n', half);
            if (cutAt >= 0)
                File.WriteAllText(FilePath, text[(cutAt + 1)..], Encoding.UTF8);
        }
        catch { }
    }
}
