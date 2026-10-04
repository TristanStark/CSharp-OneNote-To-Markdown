using System.Net;
using System.Text;
using System.Text.RegularExpressions;

namespace OneNoteMarkdown;

internal static partial class MarkdownConverter
{
    public static string ToHtml(string markdown)
    {
        var lines = NormalizeNewlines(markdown).Split('\n');
        var html = new StringBuilder();
        var paragraph = new List<string>();
        var inFence = false;
        var fenceLanguage = string.Empty;
        var fence = new StringBuilder();
        string? openList = null;

        void FlushParagraph()
        {
            if (paragraph.Count == 0)
            {
                return;
            }

            html.Append("<p>")
                .Append(ConvertInline(string.Join(" ", paragraph).Trim()))
                .AppendLine("</p>");
            paragraph.Clear();
        }

        void CloseList()
        {
            if (openList is null)
            {
                return;
            }

            html.Append("</").Append(openList).AppendLine(">");
            openList = null;
        }

        for (var i = 0; i < lines.Length; i++)
        {
            var raw = lines[i];
            var line = raw.TrimEnd();

            if (inFence)
            {
                if (line.TrimStart().StartsWith("```", StringComparison.Ordinal))
                {
                    html.Append("<pre><code");
                    if (!string.IsNullOrWhiteSpace(fenceLanguage))
                    {
                        html.Append(" data-language=\"")
                            .Append(WebUtility.HtmlEncode(fenceLanguage))
                            .Append('"');
                    }

                    html.Append('>')
                        .Append(WebUtility.HtmlEncode(fence.ToString().TrimEnd('\r', '\n')))
                        .AppendLine("</code></pre>");
                    inFence = false;
                    fenceLanguage = string.Empty;
                    fence.Clear();
                }
                else
                {
                    fence.AppendLine(raw);
                }

                continue;
            }

            if (line.TrimStart().StartsWith("```", StringComparison.Ordinal))
            {
                FlushParagraph();
                CloseList();
                inFence = true;
                fenceLanguage = line.Trim()[3..].Trim();
                continue;
            }

            if (string.IsNullOrWhiteSpace(line))
            {
                FlushParagraph();
                CloseList();
                continue;
            }

            var heading = HeadingRegex().Match(line);
            if (heading.Success)
            {
                FlushParagraph();
                CloseList();
                var level = heading.Groups[1].Value.Length;
                html.Append("<h").Append(level).Append('>')
                    .Append(ConvertInline(heading.Groups[2].Value.Trim()))
                    .Append("</h").Append(level).AppendLine(">");
                continue;
            }

            if (i + 1 < lines.Length && LooksLikeTableRow(line) && IsTableSeparator(lines[i + 1]))
            {
                FlushParagraph();
                CloseList();
                var rows = new List<string[]> { SplitPipeRow(line) };
                i += 2;
                while (i < lines.Length && LooksLikeTableRow(lines[i]) && !string.IsNullOrWhiteSpace(lines[i]))
                {
                    rows.Add(SplitPipeRow(lines[i]));
                    i++;
                }

                i--;
                AppendTable(html, rows);
                continue;
            }

            var unordered = UnorderedListRegex().Match(line);
            var ordered = OrderedListRegex().Match(line);
            if (unordered.Success || ordered.Success)
            {
                FlushParagraph();
                var listTag = unordered.Success ? "ul" : "ol";
                if (openList != listTag)
                {
                    CloseList();
                    openList = listTag;
                    html.Append('<').Append(openList).AppendLine(">");
                }

                var content = unordered.Success
                    ? unordered.Groups[1].Value
                    : ordered.Groups[1].Value;

                if (CheckboxRegex().Match(content) is { Success: true } checkbox)
                {
                    var checkedMark = checkbox.Groups[1].Value.Equals("x", StringComparison.OrdinalIgnoreCase)
                        ? "☑ "
                        : "☐ ";
                    content = checkedMark + checkbox.Groups[2].Value;
                }

                html.Append("<li>")
                    .Append(ConvertInline(content.Trim()))
                    .AppendLine("</li>");
                continue;
            }

            if (line.TrimStart().StartsWith("> ", StringComparison.Ordinal))
            {
                FlushParagraph();
                CloseList();
                html.Append("<p><em>")
                    .Append(ConvertInline(line.TrimStart()[2..]))
                    .AppendLine("</em></p>");
                continue;
            }

            if (HorizontalRuleRegex().IsMatch(line))
            {
                FlushParagraph();
                CloseList();
                html.AppendLine("<p>────────</p>");
                continue;
            }

            paragraph.Add(line.Trim());
        }

        if (inFence)
        {
            html.Append("<pre><code>")
                .Append(WebUtility.HtmlEncode(fence.ToString()))
                .AppendLine("</code></pre>");
        }

        FlushParagraph();
        CloseList();
        return html.ToString().Trim();
    }

    public static string FromHtml(string html)
    {
        if (string.IsNullOrWhiteSpace(html))
        {
            return string.Empty;
        }

        var value = NormalizeNewlines(html);

        value = Regex.Replace(value, @"<!--.*?-->", string.Empty, RegexOptions.Singleline);
        value = Regex.Replace(value, @"<(script|style|head)\b[^>]*>.*?</\1>", string.Empty, HtmlOptions);

        value = Regex.Replace(
            value,
            @"<pre\b[^>]*>\s*<code\b[^>]*>(.*?)</code>\s*</pre>",
            m => "\n\n```\n" + WebUtility.HtmlDecode(StripTags(m.Groups[1].Value)).Trim('\r', '\n') + "\n```\n\n",
            HtmlOptions);

        value = Regex.Replace(
            value,
            @"<pre\b[^>]*>(.*?)</pre>",
            m => "\n\n```\n" + WebUtility.HtmlDecode(StripTags(m.Groups[1].Value)).Trim('\r', '\n') + "\n```\n\n",
            HtmlOptions);

        value = Regex.Replace(value, @"<table\b[^>]*>.*?</table>", ConvertTable, HtmlOptions);
        value = Regex.Replace(value, @"<ol\b[^>]*>.*?</ol>", m => ConvertList(m.Value, ordered: true), HtmlOptions);
        value = Regex.Replace(value, @"<ul\b[^>]*>.*?</ul>", m => ConvertList(m.Value, ordered: false), HtmlOptions);

        value = Regex.Replace(
            value,
            @"<h([1-6])\b[^>]*>(.*?)</h\1>",
            m => "\n\n" + new string('#', int.Parse(m.Groups[1].Value)) + " " + InlineHtmlToMarkdown(m.Groups[2].Value) + "\n\n",
            HtmlOptions);

        value = Regex.Replace(
            value,
            @"<a\b[^>]*href\s*=\s*[""']([^""']+)[""'][^>]*>(.*?)</a>",
            m => $"[{InlineHtmlToMarkdown(m.Groups[2].Value)}]({WebUtility.HtmlDecode(m.Groups[1].Value)})",
            HtmlOptions);

        value = ReplacePair(value, "strong", "**");
        value = ReplacePair(value, "b", "**");
        value = ReplacePair(value, "em", "*");
        value = ReplacePair(value, "i", "*");
        value = ReplacePair(value, "del", "~~");
        value = ReplacePair(value, "s", "~~");
        value = ReplacePair(value, "code", "`");

        value = Regex.Replace(value, @"<br\s*/?>", "\n", RegexOptions.IgnoreCase);
        value = Regex.Replace(value, @"</?(p|div|section|article)\b[^>]*>", "\n\n", RegexOptions.IgnoreCase);
        value = Regex.Replace(value, @"</?(span|font|body|html)\b[^>]*>", string.Empty, RegexOptions.IgnoreCase);
        value = StripTags(value);
        value = WebUtility.HtmlDecode(value).Replace('\u00A0', ' ');

        return NormalizeMarkdownWhitespace(value);
    }

    private static string ConvertInline(string text)
    {
        var encoded = WebUtility.HtmlEncode(text);
        var placeholders = new Dictionary<string, string>();
        var placeholderIndex = 0;

        encoded = Regex.Replace(
            encoded,
            @"`([^`\r\n]+)`",
            m =>
            {
                var key = $"@@CODE{placeholderIndex++}@@";
                placeholders[key] = $"<code>{m.Groups[1].Value}</code>";
                return key;
            });

        encoded = Regex.Replace(
            encoded,
            @"!\[([^\]]*)\]\(([^)]+)\)",
            m => $"<a href=\"{WebUtility.HtmlEncode(WebUtility.HtmlDecode(m.Groups[2].Value.Trim()))}\">[image: {m.Groups[1].Value}]</a>");

        encoded = Regex.Replace(
            encoded,
            @"\[([^\]]+)\]\(([^)]+)\)",
            m => $"<a href=\"{WebUtility.HtmlEncode(WebUtility.HtmlDecode(m.Groups[2].Value.Trim()))}\">{m.Groups[1].Value}</a>");

        encoded = Regex.Replace(encoded, @"\*\*(.+?)\*\*", "<strong>$1</strong>");
        encoded = Regex.Replace(encoded, @"~~(.+?)~~", "<del>$1</del>");
        encoded = Regex.Replace(encoded, @"(?<!\*)\*([^*\r\n]+)\*(?!\*)", "<em>$1</em>");

        foreach (var (key, replacement) in placeholders)
        {
            encoded = encoded.Replace(key, replacement, StringComparison.Ordinal);
        }

        return encoded;
    }

    private static void AppendTable(StringBuilder html, IReadOnlyList<string[]> rows)
    {
        if (rows.Count == 0)
        {
            return;
        }

        html.AppendLine("<table border=\"1\">");
        foreach (var row in rows)
        {
            html.AppendLine("<tr>");
            foreach (var cell in row)
            {
                html.Append("<td>")
                    .Append(ConvertInline(cell.Trim()))
                    .AppendLine("</td>");
            }

            html.AppendLine("</tr>");
        }

        html.AppendLine("</table>");
    }

    private static string ConvertTable(Match match)
    {
        var rows = Regex.Matches(match.Value, @"<tr\b[^>]*>(.*?)</tr>", HtmlOptions)
            .Cast<Match>()
            .Select(row => Regex.Matches(row.Groups[1].Value, @"<t[dh]\b[^>]*>(.*?)</t[dh]>", HtmlOptions)
                .Cast<Match>()
                .Select(cell => SingleLine(InlineHtmlToMarkdown(cell.Groups[1].Value)))
                .ToArray())
            .Where(cells => cells.Length > 0)
            .ToList();

        if (rows.Count == 0)
        {
            return string.Empty;
        }

        var width = rows.Max(r => r.Length);
        var sb = new StringBuilder("\n\n");
        AppendMarkdownTableRow(sb, rows[0], width);
        AppendMarkdownTableRow(sb, Enumerable.Repeat("---", width).ToArray(), width);
        foreach (var row in rows.Skip(1))
        {
            AppendMarkdownTableRow(sb, row, width);
        }

        return sb.AppendLine().ToString();
    }

    private static string ConvertList(string listHtml, bool ordered)
    {
        var items = Regex.Matches(listHtml, @"<li\b[^>]*>(.*?)</li>", HtmlOptions)
            .Cast<Match>()
            .Select(m => InlineHtmlToMarkdown(m.Groups[1].Value))
            .ToList();

        if (items.Count == 0)
        {
            return string.Empty;
        }

        var sb = new StringBuilder("\n");
        for (var i = 0; i < items.Count; i++)
        {
            sb.Append(ordered ? $"{i + 1}. " : "- ")
                .AppendLine(SingleLine(items[i]));
        }

        return sb.AppendLine().ToString();
    }

    private static string InlineHtmlToMarkdown(string value)
    {
        value = Regex.Replace(
            value,
            @"<a\b[^>]*href\s*=\s*[""']([^""']+)[""'][^>]*>(.*?)</a>",
            m => $"[{InlineHtmlToMarkdown(m.Groups[2].Value)}]({WebUtility.HtmlDecode(m.Groups[1].Value)})",
            HtmlOptions);

        value = ReplacePair(value, "strong", "**");
        value = ReplacePair(value, "b", "**");
        value = ReplacePair(value, "em", "*");
        value = ReplacePair(value, "i", "*");
        value = ReplacePair(value, "del", "~~");
        value = ReplacePair(value, "s", "~~");
        value = ReplacePair(value, "code", "`");
        value = Regex.Replace(value, @"<br\s*/?>", "\n", RegexOptions.IgnoreCase);
        value = StripTags(value);
        return WebUtility.HtmlDecode(value).Replace('\u00A0', ' ').Trim();
    }

    private static string ReplacePair(string input, string tag, string marker) =>
        Regex.Replace(
            input,
            $@"<{tag}\b[^>]*>(.*?)</{tag}>",
            m => marker + m.Groups[1].Value + marker,
            HtmlOptions);

    private static string StripTags(string input) =>
        Regex.Replace(input, @"<[^>]+>", string.Empty, RegexOptions.Singleline);

    private static void AppendMarkdownTableRow(StringBuilder sb, IReadOnlyList<string> row, int width)
    {
        sb.Append('|');
        for (var i = 0; i < width; i++)
        {
            var value = i < row.Count ? row[i] : string.Empty;
            sb.Append(' ').Append(value.Replace("|", @"\|", StringComparison.Ordinal)).Append(" |");
        }

        sb.AppendLine();
    }

    private static bool LooksLikeTableRow(string line) =>
        line.Count(c => c == '|') >= 2;

    private static bool IsTableSeparator(string line)
    {
        var cells = SplitPipeRow(line);
        return cells.Length > 0 && cells.All(cell => Regex.IsMatch(cell.Trim(), @"^:?-{3,}:?$"));
    }

    private static string[] SplitPipeRow(string line)
    {
        var trimmed = line.Trim();
        if (trimmed.StartsWith('|'))
        {
            trimmed = trimmed[1..];
        }

        if (trimmed.EndsWith('|'))
        {
            trimmed = trimmed[..^1];
        }

        return Regex.Split(trimmed, @"(?<!\\)\|")
            .Select(cell => cell.Replace(@"\|", "|", StringComparison.Ordinal).Trim())
            .ToArray();
    }

    private static string NormalizeMarkdownWhitespace(string value)
    {
        var lines = NormalizeNewlines(value)
            .Split('\n')
            .Select(line => line.TrimEnd())
            .ToList();

        while (lines.Count > 0 && string.IsNullOrWhiteSpace(lines[0]))
        {
            lines.RemoveAt(0);
        }

        while (lines.Count > 0 && string.IsNullOrWhiteSpace(lines[^1]))
        {
            lines.RemoveAt(lines.Count - 1);
        }

        var output = new List<string>();
        var blank = false;
        foreach (var line in lines)
        {
            if (string.IsNullOrWhiteSpace(line))
            {
                if (!blank)
                {
                    output.Add(string.Empty);
                }

                blank = true;
            }
            else
            {
                output.Add(line);
                blank = false;
            }
        }

        return string.Join(Environment.NewLine, output).Trim();
    }

    private static string NormalizeNewlines(string value) =>
        value.Replace("\r\n", "\n", StringComparison.Ordinal)
            .Replace('\r', '\n');

    private static string SingleLine(string value) =>
        Regex.Replace(value, @"\s*\n\s*", " ").Trim();

    private static readonly RegexOptions HtmlOptions =
        RegexOptions.IgnoreCase | RegexOptions.Singleline | RegexOptions.CultureInvariant;

    [GeneratedRegex(@"^(#{1,6})\s+(.+)$")]
    private static partial Regex HeadingRegex();

    [GeneratedRegex(@"^\s*[-+*]\s+(.+)$")]
    private static partial Regex UnorderedListRegex();

    [GeneratedRegex(@"^\s*\d+[.)]\s+(.+)$")]
    private static partial Regex OrderedListRegex();

    [GeneratedRegex(@"^\s*\[([ xX])\]\s+(.+)$")]
    private static partial Regex CheckboxRegex();

    [GeneratedRegex(@"^\s*((-{3,})|(\*{3,})|(_{3,}))\s*$")]
    private static partial Regex HorizontalRuleRegex();
}
