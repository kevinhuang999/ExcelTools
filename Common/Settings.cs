namespace ExcelTools.Common;

using System.IO;
using System.Text.Json;

/// <summary>
/// Per-user settings store, persisted to %APPDATA%\ExcelTools\settings.json.
///
/// Why this exists: the original add-in hard-coded its column/sheet names, so
/// every tool ran with a single click. This rewrite intentionally contains no
/// organization-specific information, so instead each form remembers whatever
/// the user last chose and re-fills it on open. The normal flow is therefore
/// still "open the dialog, press the action button" - nothing to type - while
/// no customer-specific data lives in the source code.
///
/// Keys are plain strings; callers pass a neutral fallback.
/// </summary>
internal static class Settings
{
    private static readonly string Folder = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "ExcelTools");

    private static readonly string StorePath = Path.Combine(Folder, "settings.json");

    private static Dictionary<string, string> _values = Load();

    private static Dictionary<string, string> Load()
    {
        try
        {
            if (File.Exists(StorePath))
            {
                var json = File.ReadAllText(StorePath);
                var dict = JsonSerializer.Deserialize<Dictionary<string, string>>(json);
                if (dict != null)
                    return new Dictionary<string, string>(dict, StringComparer.OrdinalIgnoreCase);
            }
        }
        catch
        {
            // A corrupt or unreadable settings file must never break the add-in.
        }
        return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
    }

    /// <summary>Last value the user entered, or <paramref name="fallback"/>.</summary>
    public static string Get(string key, string fallback)
    {
        return _values.TryGetValue(key, out var v) && !string.IsNullOrWhiteSpace(v) ? v : fallback;
    }

    /// <summary>Remember one value and persist immediately.</summary>
    public static void Set(string key, string value)
    {
        _values[key] = value ?? string.Empty;
        Save();
    }

    /// <summary>Remember several values in one write.</summary>
    public static void SetMany(params (string Key, string Value)[] pairs)
    {
        foreach (var pair in pairs)
            _values[pair.Key] = pair.Value ?? string.Empty;
        Save();
    }

    private static void Save()
    {
        try
        {
            Directory.CreateDirectory(Folder);
            File.WriteAllText(
                StorePath,
                JsonSerializer.Serialize(_values, new JsonSerializerOptions { WriteIndented = true }));
        }
        catch
        {
            // Persisting preferences is best-effort; never surface an error.
        }
    }
}
