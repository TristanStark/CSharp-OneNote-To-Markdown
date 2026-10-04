using System.Text;
using System.Text.RegularExpressions;

namespace OneNoteMarkdown;

internal static partial class Utilities
{
    public static string SanitizeFileName(string value, string fallback = "page")
    {
        var invalid = Path.GetInvalidFileNameChars();
        var cleaned = new string(value.Select(c => invalid.Contains(c) ? '_' : c).ToArray()).Trim();

        cleaned = WhitespaceRegex().Replace(cleaned, " ");
        cleaned = cleaned.TrimEnd('.', ' ');

        return string.IsNullOrWhiteSpace(cleaned) ? fallback : cleaned;
    }

    public static string MakeUniquePath(string path)
    {
        if (!File.Exists(path) && !Directory.Exists(path))
        {
            return path;
        }

        var directory = Path.GetDirectoryName(path) ?? Environment.CurrentDirectory;
        var stem = Path.GetFileNameWithoutExtension(path);
        var extension = Path.GetExtension(path);

        for (var i = 2; ; i++)
        {
            var candidate = Path.Combine(directory, $"{stem}-{i}{extension}");
            if (!File.Exists(candidate) && !Directory.Exists(candidate))
            {
                return candidate;
            }
        }
    }

    public static string EscapeMarkdownHeading(string value) =>
        value.Replace("\r", " ", StringComparison.Ordinal)
            .Replace("\n", " ", StringComparison.Ordinal)
            .Trim();

    public static string RelativeMarkdownPath(string fromDirectory, string targetPath) =>
        Path.GetRelativePath(fromDirectory, targetPath)
            .Replace(Path.DirectorySeparatorChar, '/')
            .Replace(Path.AltDirectorySeparatorChar, '/');

    public static string DetectImageExtension(byte[] bytes, string? formatHint)
    {
        var hint = (formatHint ?? string.Empty).Trim().TrimStart('.').ToLowerInvariant();
        if (hint is "jpeg" or "jpg")
        {
            return ".jpg";
        }

        if (hint is "png")
        {
            return ".png";
        }

        if (hint is "gif")
        {
            return ".gif";
        }

        if (hint is "bmp")
        {
            return ".bmp";
        }

        if (hint is "tif" or "tiff")
        {
            return ".tiff";
        }

        if (bytes.Length >= 8 &&
            bytes[0] == 0x89 && bytes[1] == 0x50 && bytes[2] == 0x4E && bytes[3] == 0x47)
        {
            return ".png";
        }

        if (bytes.Length >= 3 && bytes[0] == 0xFF && bytes[1] == 0xD8 && bytes[2] == 0xFF)
        {
            return ".jpg";
        }

        if (bytes.Length >= 6 && Encoding.ASCII.GetString(bytes, 0, 3) == "GIF")
        {
            return ".gif";
        }

        if (bytes.Length >= 2 && bytes[0] == 0x42 && bytes[1] == 0x4D)
        {
            return ".bmp";
        }

        return ".bin";
    }

    public static string DetectOneNoteImageFormat(string path)
    {
        return Path.GetExtension(path).ToLowerInvariant() switch
        {
            ".png" => "png",
            ".jpg" or ".jpeg" => "jpg",
            ".gif" => "gif",
            ".bmp" => "bmp",
            ".tif" or ".tiff" => "tiff",
            _ => throw new NotSupportedException(
                $"Image format '{Path.GetExtension(path)}' is not supported for embedding. " +
                "Supported formats: PNG, JPEG, GIF, BMP, TIFF.")
        };
    }

    [GeneratedRegex(@"\s+")]
    private static partial Regex WhitespaceRegex();
}
