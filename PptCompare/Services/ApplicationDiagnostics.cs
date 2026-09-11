using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security;
using System.Text;
using PptCompare.Models;

namespace PptCompare.Services;

public interface IApplicationDiagnostics
{
    string DisplayLogLocation { get; }

    void RecordEvent(string eventName, string? nonSensitiveDetail = null);

    void RecordException(string eventName, Exception exception);

    string CreateSupportSummary();
}

public sealed class FileApplicationDiagnostics : IApplicationDiagnostics
{
    private const long MaxLogBytes = 1024 * 1024;
    private const int MaxArchivedLogs = 4;
    private const int MaxTokenLength = 160;
    private readonly object _writeGate = new();
    private readonly string _logDirectory;
    private readonly string _logPath;

    public FileApplicationDiagnostics()
        : this(CreateDefaultLogDirectory())
    {
    }

    internal FileApplicationDiagnostics(string logDirectory)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(logDirectory);
        _logDirectory = Path.GetFullPath(logDirectory);
        _logPath = Path.Combine(_logDirectory, "PptCompare.log");
    }

    public string DisplayLogLocation => @"%LocalAppData%\PptCompare\Logs";

    public void RecordEvent(string eventName, string? nonSensitiveDetail = null) =>
        Write("INFO", eventName, nonSensitiveDetail);

    public void RecordException(string eventName, Exception exception)
    {
        ArgumentNullException.ThrowIfNull(exception);
        var detail = $"type={SanitizeToken(exception.GetType().FullName)};hresult=0x{exception.HResult:X8}";
        if (exception.InnerException is { } innerException)
        {
            detail += $";inner={SanitizeToken(innerException.GetType().FullName)}";
        }

        Write("ERROR", eventName, detail);
    }

    public string CreateSupportSummary()
    {
        var assembly = typeof(ApplicationInfo).Assembly;
        var openXmlVersion = assembly
            .GetReferencedAssemblies()
            .FirstOrDefault(reference => string.Equals(
                reference.Name,
                "DocumentFormat.OpenXml",
                StringComparison.Ordinal))?
            .Version?
            .ToString() ?? "Unknown";
        var builder = new StringBuilder();
        builder.AppendLine(CultureInfo.InvariantCulture, $"Product: {ApplicationInfo.ProductName}");
        builder.AppendLine(CultureInfo.InvariantCulture, $"Version: {ApplicationInfo.Version}");
        builder.AppendLine(CultureInfo.InvariantCulture, $"Runtime: {RuntimeInformation.FrameworkDescription}");
        builder.AppendLine(CultureInfo.InvariantCulture, $"Operating system: {RuntimeInformation.OSDescription}");
        builder.AppendLine(CultureInfo.InvariantCulture, $"Process architecture: {RuntimeInformation.ProcessArchitecture}");
        builder.AppendLine(CultureInfo.InvariantCulture, $"Open XML SDK: {openXmlVersion}");
        builder.Append("Diagnostics policy: filenames, file paths and presentation content are not recorded.");
        return builder.ToString();
    }

    [SuppressMessage(
        "Design",
        "CA1031:Do not catch general exception types",
        Justification = "Diagnostics must never cause the application operation being recorded to fail.")]
    private void Write(string level, string eventName, string? detail)
    {
        try
        {
            var safeEventName = SanitizeToken(eventName);
            var safeDetail = string.IsNullOrWhiteSpace(detail) ? null : SanitizeDetail(detail);
            var line = string.Create(
                CultureInfo.InvariantCulture,
                $"{DateTimeOffset.UtcNow:O}|{level}|{safeEventName}{(safeDetail is null ? string.Empty : $"|{safeDetail}")}{Environment.NewLine}");
            lock (_writeGate)
            {
                Directory.CreateDirectory(_logDirectory);
                RotateIfRequired(Encoding.UTF8.GetByteCount(line));
                File.AppendAllText(_logPath, line, new UTF8Encoding(false));
            }
        }
        catch (Exception exception)
        {
            Debug.WriteLine($"PptCompare diagnostics were unavailable: {exception.GetType().Name}");
        }
    }

    private void RotateIfRequired(int incomingBytes)
    {
        if (!File.Exists(_logPath) || new FileInfo(_logPath).Length + incomingBytes <= MaxLogBytes)
        {
            return;
        }

        for (var index = MaxArchivedLogs; index >= 1; index--)
        {
            var currentPath = GetArchivePath(index);
            if (!File.Exists(currentPath))
            {
                continue;
            }

            if (index == MaxArchivedLogs)
            {
                File.Delete(currentPath);
            }
            else
            {
                File.Move(currentPath, GetArchivePath(index + 1), true);
            }
        }

        File.Move(_logPath, GetArchivePath(1), true);
    }

    private string GetArchivePath(int index) => Path.Combine(_logDirectory, $"PptCompare.{index}.log");

    private static string CreateDefaultLogDirectory()
    {
        var localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        if (string.IsNullOrWhiteSpace(localAppData))
        {
            throw new InvalidOperationException("The per-user application data folder is unavailable.");
        }

        return Path.Combine(localAppData, "PptCompare", "Logs");
    }

    private static string SanitizeDetail(string value)
    {
        var result = new StringBuilder(Math.Min(value.Length, MaxTokenLength));
        foreach (var character in value)
        {
            if (result.Length >= MaxTokenLength)
            {
                break;
            }

            if (char.IsLetterOrDigit(character) || character is ' ' or '.' or '-' or '_' or '=' or ';' or ':')
            {
                result.Append(character);
            }
        }

        return result.ToString();
    }

    private static string SanitizeToken(string? value) =>
        string.IsNullOrWhiteSpace(value)
            ? "Unknown"
            : SanitizeDetail(value);
}

internal sealed class NullApplicationDiagnostics : IApplicationDiagnostics
{
    public static NullApplicationDiagnostics Instance { get; } = new();

    public string DisplayLogLocation => @"%LocalAppData%\PptCompare\Logs";

    public void RecordEvent(string eventName, string? nonSensitiveDetail = null)
    {
    }

    public void RecordException(string eventName, Exception exception)
    {
    }

    public string CreateSupportSummary() => string.Empty;
}
