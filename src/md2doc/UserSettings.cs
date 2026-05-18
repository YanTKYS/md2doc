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

    // 変換エンジン: "OpenXml"（既定）または "WordCom"
    // 既存設定に項目がない場合は既定値 "OpenXml" が使われる（JSON デシリアライズのデフォルト動作）
    public string ConversionEngine { get; set; } = "OpenXml";

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
            var settings = JsonSerializer.Deserialize<UserSettings>(json) ?? new();
            settings.Normalize();
            return settings;
        }
        catch { return new(); }
    }

    // 設定ファイルが手動編集された場合や、過去バージョンとの非互換に備え、
    // 不正な値を既定値に丸めて UI 側の switch default 等で誤った挙動を防ぐ。
    private void Normalize()
    {
        if (string.IsNullOrWhiteSpace(BodyFontName)) BodyFontName = "MS Gothic";
        BodyFontSize = Math.Clamp(BodyFontSize, 8.0, 72.0);
        if (HeaderMode is < 0 or > 2) HeaderMode = 0;
        HeaderCustomText ??= "";
        if (HeaderAlignment is < 0 or > 2) HeaderAlignment = 0;
        if (FooterAlignment is < 0 or > 2) FooterAlignment = 1;
        if (ConversionEngine != "OpenXml" && ConversionEngine != "WordCom") ConversionEngine = "OpenXml";
        LastOutputFolder ??= "";
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
