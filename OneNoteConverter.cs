using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace OneNoteMarkdown;

internal static partial class OneNoteConverter
{
    private const string OneNoteNamespace = "http://schemas.microsoft.com/office/onenote/2013/onenote";

    public static void ImportMarkdown(string input, string outputOne)
    {
        if (!OperatingSystem.IsWindows())
        {
            throw new PlatformNotSupportedException("Import requires Windows and Microsoft OneNote Desktop.");
        }

        var markdownFiles = ResolveMarkdownFiles(input);
        if (markdownFiles.Count == 0)
        {
            throw new InvalidOperationException("No Markdown files were found.");
        }

        if (!outputOne.EndsWith(".one", StringComparison.OrdinalIgnoreCase))
        {
            throw new ArgumentException("The output path must end in .one.");
        }

        if (File.Exists(outputOne))
        {
            throw new IOException(
                $"Output section already exists: {outputOne}. Refusing to overwrite an existing OneNote section.");
        }

        Directory.CreateDirectory(Path.GetDirectoryName(outputOne) ?? Environment.CurrentDirectory);

        using var oneNote = new OneNoteComClient();
        var sectionId = oneNote.OpenSection(outputOne, create: true);

        foreach (var markdownFile in markdownFiles)
        {
            var markdown = File.ReadAllText(markdownFile, Encoding.UTF8);
            var (title, body) = ExtractTitle(markdown, markdownFile);
            var pageId = oneNote.CreatePage(sectionId);
            var blocks = BuildImportBlocks(
                body,
                Path.GetDirectoryName(markdownFile) ?? Environment.CurrentDirectory);

            oneNote.UpdatePageContent(BuildPageXml(pageId, title, blocks));
        }
    }

    public static IReadOnlyList<string> ExportOneNote(string inputOne, string output)
    {
        if (!OperatingSystem.IsWindows())
        {
            throw new PlatformNotSupportedException("Export requires Windows and Microsoft OneNote Desktop.");
        }

        if (!File.Exists(inputOne))
        {
            throw new FileNotFoundException("OneNote section not found.", inputOne);
        }

        using var oneNote = new OneNoteComClient();
        var sectionId = oneNote.OpenSection(inputOne, create: false);
        var pages = ReadPages(oneNote.GetPageHierarchy(sectionId));

        if (pages.Count == 0)
        {
            throw new InvalidOperationException("The OneNote section contains no pages.");
        }

        var singleFile = output.EndsWith(".md", StringComparison.OrdinalIgnoreCase);
        if (singleFile && pages.Count != 1)
        {
            throw new InvalidOperationException(
                $"The section contains {pages.Count} pages. Export to a directory instead of a single .md file.");
        }

        var result = new List<string>();

        if (singleFile)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(output) ?? Environment.CurrentDirectory);
            ExportPage(oneNote, pages[0], output);
            result.Add(output);
            return result;
        }

        Directory.CreateDirectory(output);

        for (var i = 0; i < pages.Count; i++)
        {
            var safeTitle = Utilities.SanitizeFileName(pages[i].Name, $"page-{i + 1}");
            var target = Utilities.MakeUniquePath(
                Path.Combine(output, $"{i + 1:000} - {safeTitle}.md"));

            ExportPage(oneNote, pages[i], target);
            result.Add(target);
        }

        return result;
    }

    private static void ExportPage(OneNoteComClient oneNote, PageRef page, string markdownPath)
    {
        var pageXml = oneNote.GetPageContent(page.Id);
        var assetsDirectory = Path.Combine(
            Path.GetDirectoryName(markdownPath) ?? Environment.CurrentDirectory,
            Path.GetFileNameWithoutExtension(markdownPath) + ".assets");

        var markdown = ConvertPageXmlToMarkdown(
            oneNote,
            page,
            pageXml,
            markdownPath,
            assetsDirectory);

        File.WriteAllText(
            markdownPath,
            markdown + Environment.NewLine,
            new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));

        if (Directory.Exists(assetsDirectory) &&
            !Directory.EnumerateFileSystemEntries(assetsDirectory).Any())
        {
            Directory.Delete(assetsDirectory);
        }
    }

    private static string ConvertPageXmlToMarkdown(
        OneNoteComClient oneNote,
        PageRef page,
        string pageXml,
        string markdownPath,
        string assetsDirectory)
    {
        var document = XDocument.Parse(pageXml, LoadOptions.PreserveWhitespace);
        var root = document.Root ?? throw new InvalidOperationException("Invalid OneNote page XML.");
        var ns = root.Name.Namespace;

        var title = root
            .Descendants(ns + "Title")
            .Descendants(ns + "T")
            .Select(t => MarkdownConverter.FromHtml(t.Value))
            .FirstOrDefault(t => !string.IsNullOrWhiteSpace(t));

        title = string.IsNullOrWhiteSpace(title) ? page.Name : title;

        var output = new StringBuilder();
        output.Append("# ")
            .AppendLine(Utilities.EscapeMarkdownHeading(title!))
            .AppendLine();

        var imageCounter = 0;

        var outlines = root.Elements(ns + "Outline")
            .OrderBy(outline => ReadPosition(outline, ns).Y)
            .ThenBy(outline => ReadPosition(outline, ns).X);

        foreach (var outline in outlines)
        {
            var children = outline.Element(ns + "OEChildren");
            if (children is null)
            {
                continue;
            }

            foreach (var child in children.Elements())
            {
                AppendNode(
                    child,
                    0,
                    output,
                    oneNote,
                    page.Id,
                    ns,
                    markdownPath,
                    assetsDirectory,
                    ref imageCounter);
            }

            EnsureBlankLine(output);
        }

        return NormalizeMarkdown(output.ToString());
    }

    private static void AppendNode(
        XElement node,
        int depth,
        StringBuilder output,
        OneNoteComClient oneNote,
        string pageId,
        XNamespace ns,
        string markdownPath,
        string assetsDirectory,
        ref int imageCounter)
    {
        if (node.Name.LocalName == "HTMLBlock")
        {
            AppendBlock(output, MarkdownConverter.FromHtml(node.Element(ns + "Data")?.Value ?? string.Empty));
            return;
        }

        if (node.Name.LocalName != "OE")
        {
            foreach (var child in node.Elements())
            {
                AppendNode(
                    child,
                    depth,
                    output,
                    oneNote,
                    pageId,
                    ns,
                    markdownPath,
                    assetsDirectory,
                    ref imageCounter);
            }

            return;
        }

        var table = node.Element(ns + "Table");
        if (table is not null)
        {
            AppendBlock(output, ConvertOneNoteTable(table, ns));
        }

        var image = node.Element(ns + "Image");
        if (image is not null)
        {
            output.AppendLine(
                    ExportImage(
                        image,
                        oneNote,
                        pageId,
                        markdownPath,
                        assetsDirectory,
                        ref imageCounter,
                        ns))
                .AppendLine();
        }

        var text = node.Element(ns + "T")?.Value;
        if (!string.IsNullOrWhiteSpace(text))
        {
            var markdown = MarkdownConverter.FromHtml(text);
            var list = node.Element(ns + "List");
            var isBullet = list?.Descendants(ns + "Bullet").Any() == true;
            var isNumber = list?.Descendants(ns + "Number").Any() == true;

            if (isBullet || isNumber)
            {
                output.Append(new string(' ', depth * 2))
                    .Append(isNumber ? "1. " : "- ")
                    .AppendLine(Regex.Replace(markdown, @"\s*\n\s*", " ").Trim());
            }
            else
            {
                AppendBlock(output, markdown);
            }
        }

        var nested = node.Element(ns + "OEChildren");
        if (nested is null)
        {
            return;
        }

        foreach (var child in nested.Elements())
        {
            AppendNode(
                child,
                depth + 1,
                output,
                oneNote,
                pageId,
                ns,
                markdownPath,
                assetsDirectory,
                ref imageCounter);
        }
    }

    private static string ExportImage(
        XElement image,
        OneNoteComClient oneNote,
        string pageId,
        string markdownPath,
        string assetsDirectory,
        ref int counter,
        XNamespace ns)
    {
        byte[]? bytes = null;
        var data = image.Element(ns + "Data")?.Value;

        if (!string.IsNullOrWhiteSpace(data))
        {
            try
            {
                bytes = Convert.FromBase64String(data);
            }
            catch (FormatException)
            {
                bytes = null;
            }
        }

        if (bytes is null)
        {
            var callbackId = image.Attribute("callbackID")?.Value;
            if (!string.IsNullOrWhiteSpace(callbackId))
            {
                var base64 = oneNote.GetBinaryPageContent(pageId, callbackId);
                if (!string.IsNullOrWhiteSpace(base64))
                {
                    bytes = Convert.FromBase64String(base64);
                }
            }
        }

        if (bytes is null || bytes.Length == 0)
        {
            return "> [Image omitted: OneNote did not expose its binary content]";
        }

        Directory.CreateDirectory(assetsDirectory);
        counter++;

        var extension = Utilities.DetectImageExtension(
            bytes,
            image.Attribute("format")?.Value);

        var target = Path.Combine(
            assetsDirectory,
            $"image-{counter:000}{extension}");

        File.WriteAllBytes(target, bytes);

        var mdDirectory =
            Path.GetDirectoryName(markdownPath) ?? Environment.CurrentDirectory;

        var relative = Utilities.RelativeMarkdownPath(mdDirectory, target);
        return $"![OneNote image]({relative})";
    }

    private static string ConvertOneNoteTable(XElement table, XNamespace ns)
    {
        var rows = table.Elements(ns + "Row")
            .Select(row => row.Elements(ns + "Cell")
                .Select(cell =>
                {
                    var text = string.Join(
                        " ",
                        cell.Descendants(ns + "T")
                            .Select(t => MarkdownConverter.FromHtml(t.Value))
                            .Where(t => !string.IsNullOrWhiteSpace(t)));

                    return Regex.Replace(text, @"\s*\n\s*", " ").Trim();
                })
                .ToArray())
            .Where(row => row.Length > 0)
            .ToList();

        if (rows.Count == 0)
        {
            return string.Empty;
        }

        var width = rows.Max(row => row.Length);
        var result = new StringBuilder();

        AppendTableRow(result, rows[0], width);
        AppendTableRow(result, Enumerable.Repeat("---", width).ToArray(), width);

        foreach (var row in rows.Skip(1))
        {
            AppendTableRow(result, row, width);
        }

        return result.ToString().TrimEnd();
    }

    private static void AppendTableRow(
        StringBuilder output,
        IReadOnlyList<string> row,
        int width)
    {
        output.Append('|');

        for (var i = 0; i < width; i++)
        {
            var value = i < row.Count ? row[i] : string.Empty;
            output.Append(' ')
                .Append(value.Replace("|", @"\|", StringComparison.Ordinal))
                .Append(" |");
        }

        output.AppendLine();
    }

    private static IReadOnlyList<ImportBlock> BuildImportBlocks(
        string markdown,
        string sourceDirectory)
    {
        var lines = markdown
            .Replace("\r\n", "\n", StringComparison.Ordinal)
            .Replace('\r', '\n')
            .Split('\n');

        var blocks = new List<ImportBlock>();
        var text = new StringBuilder();

        void FlushText()
        {
            var value = text.ToString().Trim();
            if (value.Length > 0)
            {
                blocks.Add(new HtmlImportBlock(MarkdownConverter.ToHtml(value)));
            }

            text.Clear();
        }

        foreach (var line in lines)
        {
            var match = StandaloneImageRegex().Match(line);
            if (!match.Success)
            {
                text.AppendLine(line);
                continue;
            }

            var alt = match.Groups[1].Value.Trim();
            var target = match.Groups[2].Value.Trim().Trim('<', '>');
            target = Regex.Replace(target, @"\s+[""'].*[""']\s*$", string.Empty).Trim();

            if (Uri.TryCreate(target, UriKind.Absolute, out var remote) &&
                remote.Scheme is "http" or "https")
            {
                text.AppendLine($"[{(string.IsNullOrWhiteSpace(alt) ? "image" : alt)}]({target})");
                continue;
            }

            var localPath = Path.GetFullPath(
                Path.Combine(
                    sourceDirectory,
                    Uri.UnescapeDataString(
                        target.Replace('/', Path.DirectorySeparatorChar))));

            if (!File.Exists(localPath))
            {
                text.AppendLine($"> Missing image: {target}");
                continue;
            }

            FlushText();
            blocks.Add(new ImageImportBlock(localPath));
        }

        FlushText();

        if (blocks.Count == 0)
        {
            blocks.Add(new HtmlImportBlock("<p></p>"));
        }

        return blocks;
    }

    private static string BuildPageXml(
        string pageId,
        string title,
        IReadOnlyList<ImportBlock> blocks)
    {
        XNamespace one = OneNoteNamespace;
        var children = new XElement(one + "OEChildren");

        foreach (var block in blocks)
        {
            if (block is HtmlImportBlock html)
            {
                children.Add(
                    new XElement(
                        one + "HTMLBlock",
                        new XElement(
                            one + "Data",
                            new XCData($"<html><body>{html.Html}</body></html>"))));

                continue;
            }

            if (block is ImageImportBlock image)
            {
                var bytes = File.ReadAllBytes(image.Path);
                children.Add(
                    new XElement(
                        one + "OE",
                        new XElement(
                            one + "Image",
                            new XAttribute(
                                "format",
                                Utilities.DetectOneNoteImageFormat(image.Path)),
                            new XElement(
                                one + "Data",
                                Convert.ToBase64String(bytes)))));
            }
        }

        var page = new XElement(
            one + "Page",
            new XAttribute(XNamespace.Xmlns + "one", OneNoteNamespace),
            new XAttribute("ID", pageId),
            new XElement(
                one + "Title",
                new XAttribute(
                    "style",
                    "font-family:Calibri;font-size:17.0pt"),
                new XElement(
                    one + "OE",
                    new XAttribute("alignment", "left"),
                    new XElement(
                        one + "T",
                        new XCData(title)))),
            new XElement(
                one + "Outline",
                new XElement(
                    one + "Position",
                    new XAttribute("x", "36"),
                    new XAttribute("y", "90")),
                children));

        return new XDocument(
                new XDeclaration("1.0", "utf-8", null),
                page)
            .ToString(SaveOptions.DisableFormatting);
    }

    private static (string Title, string Body) ExtractTitle(
        string markdown,
        string sourcePath)
    {
        var normalized = markdown
            .Replace("\r\n", "\n", StringComparison.Ordinal)
            .Replace('\r', '\n');

        var lines = normalized.Split('\n').ToList();

        for (var i = 0; i < lines.Count; i++)
        {
            if (string.IsNullOrWhiteSpace(lines[i]))
            {
                continue;
            }

            var match = TopLevelHeadingRegex().Match(lines[i]);
            if (match.Success)
            {
                var title = match.Groups[1].Value.Trim();
                lines.RemoveAt(i);

                return (
                    string.IsNullOrWhiteSpace(title)
                        ? Path.GetFileNameWithoutExtension(sourcePath)
                        : title,
                    string.Join('\n', lines).Trim());
            }

            break;
        }

        return (
            Path.GetFileNameWithoutExtension(sourcePath),
            normalized.Trim());
    }

    private static IReadOnlyList<string> ResolveMarkdownFiles(string input)
    {
        if (File.Exists(input))
        {
            if (!input.EndsWith(".md", StringComparison.OrdinalIgnoreCase))
            {
                throw new ArgumentException("Input file must end in .md.");
            }

            return [Path.GetFullPath(input)];
        }

        if (Directory.Exists(input))
        {
            return Directory.GetFiles(
                    input,
                    "*.md",
                    SearchOption.AllDirectories)
                .OrderBy(
                    path => path,
                    StringComparer.OrdinalIgnoreCase)
                .Select(Path.GetFullPath)
                .ToList();
        }

        throw new FileNotFoundException(
            "Markdown input was not found.",
            input);
    }

    private static IReadOnlyList<PageRef> ReadPages(string hierarchyXml)
    {
        var document = XDocument.Parse(hierarchyXml);
        var root =
            document.Root ??
            throw new InvalidOperationException(
                "Invalid OneNote hierarchy XML.");

        var ns = root.Name.Namespace;

        return root.Descendants(ns + "Page")
            .Select(page => new PageRef(
                page.Attribute("ID")?.Value ?? string.Empty,
                page.Attribute("name")?.Value ?? "Untitled"))
            .Where(page => page.Id.Length > 0)
            .ToList();
    }

    private static (double X, double Y) ReadPosition(
        XElement outline,
        XNamespace ns)
    {
        var position = outline.Element(ns + "Position");
        if (position is null)
        {
            return (double.MaxValue, double.MaxValue);
        }

        _ = double.TryParse(
            position.Attribute("x")?.Value,
            System.Globalization.NumberStyles.Float,
            System.Globalization.CultureInfo.InvariantCulture,
            out var x);

        _ = double.TryParse(
            position.Attribute("y")?.Value,
            System.Globalization.NumberStyles.Float,
            System.Globalization.CultureInfo.InvariantCulture,
            out var y);

        return (x, y);
    }

    private static void AppendBlock(
        StringBuilder output,
        string value)
    {
        value = value.Trim();
        if (value.Length == 0)
        {
            return;
        }

        output.AppendLine(value);
        EnsureBlankLine(output);
    }

    private static void EnsureBlankLine(StringBuilder output)
    {
        if (output.Length == 0)
        {
            return;
        }

        var text = output.ToString();

        if (!text.EndsWith(
                Environment.NewLine + Environment.NewLine,
                StringComparison.Ordinal))
        {
            if (!text.EndsWith(
                    Environment.NewLine,
                    StringComparison.Ordinal))
            {
                output.AppendLine();
            }

            output.AppendLine();
        }
    }

    private static string NormalizeMarkdown(string value)
    {
        value = value
            .Replace("\r\n", "\n", StringComparison.Ordinal)
            .Replace('\r', '\n');

        value = Regex.Replace(value, @"[ \t]+\n", "\n");
        value = Regex.Replace(value, @"\n{3,}", "\n\n");

        return value.Trim()
            .Replace(
                "\n",
                Environment.NewLine,
                StringComparison.Ordinal);
    }

    private sealed record PageRef(string Id, string Name);

    private abstract record ImportBlock;
    private sealed record HtmlImportBlock(string Html) : ImportBlock;
    private sealed record ImageImportBlock(string Path) : ImportBlock;

    [GeneratedRegex(@"^\s*!\[([^\]]*)\]\((.+)\)\s*$")]
    private static partial Regex StandaloneImageRegex();

    [GeneratedRegex(@"^\s*#\s+(.+?)\s*$")]
    private static partial Regex TopLevelHeadingRegex();
}
