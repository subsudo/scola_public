using System.IO;
using System.Text.Json;
using VerlaufsakteApp.Models;

namespace VerlaufsakteApp.Services;

public class UserPrefsService
{
    private readonly string _prefsPath;
    private readonly JsonSerializerOptions _jsonOptions = new()
    {
        WriteIndented = true,
        PropertyNameCaseInsensitive = true
    };

    public UserPrefsService(string prefsPath)
    {
        _prefsPath = prefsPath;
    }

    public UserPrefs Load()
    {
        try
        {
            if (!File.Exists(_prefsPath))
            {
                return new UserPrefs();
            }

            var json = File.ReadAllText(_prefsPath);
            return JsonSerializer.Deserialize<UserPrefs>(json, _jsonOptions) ?? new UserPrefs();
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"UserPrefs konnten nicht geladen werden: {ex.Message}");
            return new UserPrefs();
        }
    }

    public void Save(UserPrefs prefs)
    {
        var directory = Path.GetDirectoryName(_prefsPath);
        if (!string.IsNullOrWhiteSpace(directory))
        {
            Directory.CreateDirectory(directory);
        }

        var json = JsonSerializer.Serialize(prefs, _jsonOptions);
        File.WriteAllText(_prefsPath, json);
    }
}
