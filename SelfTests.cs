using System.Xml.Linq;

namespace OneNoteMarkdown;

internal static class SelfTests
{
    public static int Run()
    {
        var failures = new List<string>();

        Check(
            "Markdown -> HTML headings",
            MarkdownConverter.ToHtml("# Hello").Contains("<h1>Hello</h1>", StringComparison.Ordinal),
            failures);

        Check(
            "Markdown -> HTML formatting",
            MarkdownConverter.ToHtml("This is **bold** and *italic*.").Contains("<strong>bold</strong>", StringComparison.Ordinal),
            failures);

        var tableHtml = MarkdownConverter.ToHtml(
            """
            | A | B |
            |---|---|
            | 1 | 2 |
            """);

        Check("Markdown -> HTML table", tableHtml.Contains("<table", StringComparison.Ordinal), failures);

        var markdown = MarkdownConverter.FromHtml("<p>Hello <strong>world</strong></p><ul><li>One</li><li>Two</li></ul>");
        Check("HTML -> Markdown bold", markdown.Contains("**world**", StringComparison.Ordinal), failures);
        Check("HTML -> Markdown list", markdown.Contains("- One", StringComparison.Ordinal), failures);

        var pageXml = BuildMinimalPageForXmlTest();
        try
        {
            _ = XDocument.Parse(pageXml);
        }
        catch (Exception ex)
        {
            failures.Add("Generated XML is invalid: " + ex.Message);
        }

        if (failures.Count == 0)
        {
            Console.WriteLine("Self-test: PASS");
            return 0;
        }

        Console.Error.WriteLine("Self-test: FAIL");
        foreach (var failure in failures)
        {
            Console.Error.WriteLine(" - " + failure);
        }

        return 1;
    }

    private static string BuildMinimalPageForXmlTest()
    {
        XNamespace one = "http://schemas.microsoft.com/office/onenote/2013/onenote";
        var page = new XElement(
            one + "Page",
            new XAttribute(XNamespace.Xmlns + "one", one.NamespaceName),
            new XAttribute("ID", "{TEST}"),
            new XElement(
                one + "Title",
                new XElement(one + "OE", new XElement(one + "T", new XCData("Test")))),
            new XElement(
                one + "Outline",
                new XElement(
                    one + "OEChildren",
                    new XElement(
                        one + "HTMLBlock",
                        new XElement(one + "Data", new XCData("<html><body><p>Test</p></body></html>"))))));

        return new XDocument(page).ToString(SaveOptions.DisableFormatting);
    }

    private static void Check(string name, bool condition, ICollection<string> failures)
    {
        if (!condition)
        {
            failures.Add(name);
        }
    }
}
