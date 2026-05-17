using System.Text;
using System.Text.Json;

namespace Md2Doc;

internal sealed class UserSettings
{
    // フォント（本文）
    public string BodyFontName { get; set; } = "MS Gothic";
    public double BodyFontSize { get; set; } = 11.0;

    // 見出しオプション
    public bool HeadingNumbering { get; set; } = false;

    // ヘッダー: Mode 0=なし 1=ファイル名 2=自由記入
    public int HeaderMode { get; set; } = 0;
    public string HeaderCustomText { get; set; } = "";
    public int HeaderAlignment { get; set; } = 0; // 0=左 1=中央 2=右

    // フッター
    public bool FooterPageNumber { get; set; } = false;
    public int FooterAlignment { get; set; } = 1; // 0=左 1=中央 2=右

    // 出力先・ウィンドウ
    public string LastOutputFolder { get; set; } = "";
    public int WindowWidth { get; set; } = 960;
    public int WindowHeight { get; set; } = 860;

    private static readonly string FilePath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "md2doc", "settings.json");

    public static UserSettings Load()
    {
        try
        {
            if (!File.Exists(FilePath)) return new();
            var json = File.ReadAllText(FilePath, Encoding.UTF8);
            return JsonSerializer.Deserialize<UserSettings>(json) ?? new();
        }
        catch { return new(); }
    }

    public void Save()
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(FilePath)!);
            File.WriteAllText(FilePath,
                JsonSerializer.Serialize(this, new JsonSerializerOptions { WriteIndented = true }),
                Encoding.UTF8);
        }
        catch { }
    }
}
