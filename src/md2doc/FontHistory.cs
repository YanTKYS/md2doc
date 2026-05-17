using System.Text;
using System.Text.Json;

namespace Md2Doc;

internal static class FontHistory
{
    private static readonly string FilePath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "md2doc", "font_history.json");

    private const int MaxCount = 5;

    public static IReadOnlyList<string> Load()
    {
        try
        {
            if (!File.Exists(FilePath)) return [];
            var json = File.ReadAllText(FilePath, Encoding.UTF8);
            return JsonSerializer.Deserialize<List<string>>(json) ?? [];
        }
        catch
        {
            return [];
        }
    }

    public static IReadOnlyList<string> Update(params string[] fontNames)
    {
        var current = Load().ToList();
        // 重複を除去した後、逆順に処理することで最初の引数が先頭になる
        foreach (var name in fontNames.Distinct().Reverse())
        {
            current.RemoveAll(f => f == name);
            current.Insert(0, name);
        }
        var result = current.Take(MaxCount).ToList();
        Save(result);
        return result;
    }

    private static void Save(IReadOnlyList<string> history)
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(FilePath)!);
            File.WriteAllText(FilePath, JsonSerializer.Serialize(history), Encoding.UTF8);
        }
        catch
        {
            // 履歴の保存失敗は変換結果に影響しない
        }
    }
}
