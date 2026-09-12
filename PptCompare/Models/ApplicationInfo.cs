using System.Reflection;

namespace PptCompare.Models;

public static class ApplicationInfo
{
    private static readonly Lazy<string> ResolvedVersion = new(ResolveVersion);

    public const string ProductName = "PptCompare";
    public const string Designer = "Ben Davies";
    public const string Copyright = "© 2026 Ben Davies";

    public static string Version => ResolvedVersion.Value;

    private static string ResolveVersion()
    {
        var assembly = typeof(ApplicationInfo).Assembly;
        var informationalVersion = assembly
            .GetCustomAttribute<AssemblyInformationalVersionAttribute>()?
            .InformationalVersion;
        if (!string.IsNullOrWhiteSpace(informationalVersion))
        {
            var metadataSeparator = informationalVersion.IndexOf('+', StringComparison.Ordinal);
            return metadataSeparator >= 0
                ? informationalVersion[..metadataSeparator]
                : informationalVersion;
        }

        return assembly.GetName().Version?.ToString(3) ?? "0.0.0";
    }
}
