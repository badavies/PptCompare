using System.Diagnostics;
using System.IO;
using System.Security;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Windows;
using PptCompare.Models;

namespace PptCompare.Services;

public sealed class SettingsPersistenceException(string message, Exception innerException)
    : Exception(message, innerException);

public interface IApplicationSettingsService
{
    ApplicationSettings Load();

    void Save(ApplicationSettings settings);
}

public sealed class JsonApplicationSettingsService : IApplicationSettingsService
{
    private const long MaxSettingsFileBytes = 64 * 1024;
    private static readonly JsonSerializerOptions SerializerOptions = new()
    {
        PropertyNameCaseInsensitive = false,
        UnmappedMemberHandling = JsonUnmappedMemberHandling.Disallow,
        WriteIndented = true,
        MaxDepth = 16
    };

    private readonly string _settingsDirectory;
    private readonly string _settingsPath;
    private readonly IApplicationDiagnostics _diagnostics;

    public JsonApplicationSettingsService(IApplicationDiagnostics? diagnostics = null)
    {
        _diagnostics = diagnostics ?? NullApplicationDiagnostics.Instance;
        var localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        if (string.IsNullOrWhiteSpace(localAppData))
        {
            throw new InvalidOperationException("The per-user application data folder is unavailable.");
        }

        _settingsDirectory = Path.Combine(localAppData, "PptCompare");
        _settingsPath = Path.Combine(_settingsDirectory, "settings.json");
    }

    public ApplicationSettings Load()
    {
        try
        {
            var file = new FileInfo(_settingsPath);
            if (!file.Exists)
            {
                return new ApplicationSettings();
            }

            if (file.Length is <= 0 or > MaxSettingsFileBytes)
            {
                Debug.WriteLine("PptCompare settings were ignored because the file size was invalid.");
                _diagnostics.RecordEvent("SettingsReset", "reason=invalid-size");
                return new ApplicationSettings();
            }

            using var stream = new FileStream(
                _settingsPath,
                FileMode.Open,
                FileAccess.Read,
                FileShare.Read,
                16 * 1024,
                FileOptions.SequentialScan);
            var settings = JsonSerializer.Deserialize<ApplicationSettings>(stream, SerializerOptions);
            if (ApplicationSettings.TryValidate(settings, out var validationError))
            {
                return settings!;
            }

            Debug.WriteLine($"PptCompare settings were ignored: {validationError}");
            _diagnostics.RecordEvent("SettingsReset", "reason=invalid-values");
            return new ApplicationSettings();
        }
        catch (Exception exception) when (IsExpectedPersistenceException(exception))
        {
            Debug.WriteLine($"PptCompare settings could not be loaded: {exception.GetType().Name}");
            _diagnostics.RecordException("SettingsLoadFailed", exception);
            return new ApplicationSettings();
        }
    }

    public void Save(ApplicationSettings settings)
    {
        ArgumentNullException.ThrowIfNull(settings);
        if (!ApplicationSettings.TryValidate(settings, out var validationError))
        {
            throw new ArgumentException(validationError, nameof(settings));
        }

        string? temporaryPath = null;
        try
        {
            Directory.CreateDirectory(_settingsDirectory);
            temporaryPath = Path.Combine(_settingsDirectory, $"settings-{Guid.NewGuid():N}.tmp");
            using (var stream = new FileStream(
                       temporaryPath,
                       FileMode.CreateNew,
                       FileAccess.Write,
                       FileShare.None,
                       16 * 1024,
                       FileOptions.WriteThrough))
            {
                JsonSerializer.Serialize(stream, settings, SerializerOptions);
                stream.Flush(true);
                if (stream.Length > MaxSettingsFileBytes)
                {
                    throw new IOException("The settings data exceeded its safe size limit.");
                }
            }

            File.Move(temporaryPath, _settingsPath, true);
            temporaryPath = null;
        }
        catch (Exception exception) when (IsExpectedPersistenceException(exception))
        {
            _diagnostics.RecordException("SettingsSaveFailed", exception);
            throw new SettingsPersistenceException(
                "The settings could not be saved. Check that your user profile is available and try again.",
                exception);
        }
        finally
        {
            if (temporaryPath is not null)
            {
                TryDeleteTemporaryFile(temporaryPath, _diagnostics);
            }
        }
    }

    private static bool IsExpectedPersistenceException(Exception exception) =>
        exception is IOException or UnauthorizedAccessException or JsonException or NotSupportedException or
            SecurityException or ArgumentException;

    private static void TryDeleteTemporaryFile(string path, IApplicationDiagnostics diagnostics)
    {
        try
        {
            File.Delete(path);
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException or SecurityException)
        {
            Debug.WriteLine($"A temporary PptCompare settings file could not be removed: {exception.GetType().Name}");
            diagnostics.RecordException("SettingsTemporaryFileCleanupFailed", exception);
        }
    }
}

public interface ISettingsDialogService
{
    ApplicationSettings? EditSettings(ApplicationSettings currentSettings);
}

public sealed class WpfSettingsDialogService : ISettingsDialogService
{
    public ApplicationSettings? EditSettings(ApplicationSettings currentSettings)
    {
        ArgumentNullException.ThrowIfNull(currentSettings);
        var window = new SettingsWindow(currentSettings)
        {
            Owner = Application.Current?.MainWindow
        };
        return window.ShowDialog() == true ? window.Result : null;
    }
}
