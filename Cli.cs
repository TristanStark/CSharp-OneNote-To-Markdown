namespace OneNoteMarkdown;

internal static class Cli
{
    public static int Run(string[] args)
    {
        if (args.Length == 0 || IsHelp(args[0]))
        {
            PrintHelp();
            return 0;
        }

        var command = args[0].ToLowerInvariant();
        return command switch
        {
            "import" => Import(args),
            "export" => Export(args),
            "convert" => Convert(args),
            "selftest" => SelfTests.Run(),
            _ => Unknown(command)
        };
    }

    private static int Import(string[] args)
    {
        if (args.Length is < 2 or > 3)
        {
            Console.Error.WriteLine("Usage: onenotemd import <input.md|directory> [output.one]");
            return 2;
        }

        var input = Path.GetFullPath(args[1]);
        var output = args.Length == 3
            ? Path.GetFullPath(args[2])
            : GetDefaultOneNoteOutput(input);

        OneNoteConverter.ImportMarkdown(input, output);
        Console.WriteLine($"Created: {output}");
        return 0;
    }

    private static int Export(string[] args)
    {
        if (args.Length is < 2 or > 3)
        {
            Console.Error.WriteLine("Usage: onenotemd export <input.one> [output.md|directory]");
            return 2;
        }

        var input = Path.GetFullPath(args[1]);
        var output = args.Length == 3
            ? Path.GetFullPath(args[2])
            : Path.Combine(
                Path.GetDirectoryName(input) ?? Environment.CurrentDirectory,
                Path.GetFileNameWithoutExtension(input) + "-markdown");

        var exported = OneNoteConverter.ExportOneNote(input, output);
        foreach (var file in exported)
        {
            Console.WriteLine($"Created: {file}");
        }

        return 0;
    }

    private static int Convert(string[] args)
    {
        if (args.Length is < 2 or > 3)
        {
            Console.Error.WriteLine("Usage: onenotemd convert <input.md|input.one|directory> [output]");
            return 2;
        }

        var input = Path.GetFullPath(args[1]);
        var output = args.Length == 3 ? Path.GetFullPath(args[2]) : null;

        if (Directory.Exists(input) || input.EndsWith(".md", StringComparison.OrdinalIgnoreCase))
        {
            output ??= GetDefaultOneNoteOutput(input);
            OneNoteConverter.ImportMarkdown(input, output);
            Console.WriteLine($"Created: {output}");
            return 0;
        }

        if (input.EndsWith(".one", StringComparison.OrdinalIgnoreCase))
        {
            output ??= Path.Combine(
                Path.GetDirectoryName(input) ?? Environment.CurrentDirectory,
                Path.GetFileNameWithoutExtension(input) + "-markdown");

            foreach (var file in OneNoteConverter.ExportOneNote(input, output))
            {
                Console.WriteLine($"Created: {file}");
            }

            return 0;
        }

        throw new ArgumentException("Input must be a .md file, a directory containing Markdown files, or a .one section.");
    }

    private static string GetDefaultOneNoteOutput(string input)
    {
        if (Directory.Exists(input))
        {
            return input.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar) + ".one";
        }

        return Path.ChangeExtension(input, ".one");
    }

    private static bool IsHelp(string value) =>
        value is "-h" or "--help" or "help" or "/?";

    private static int Unknown(string command)
    {
        Console.Error.WriteLine($"Unknown command: {command}");
        PrintHelp();
        return 2;
    }

    private static void PrintHelp()
    {
        Console.WriteLine(
            """
            onenotemd - Markdown <-> Microsoft OneNote (.one), without third-party runtime dependencies

            Requirements:
              - Windows
              - Microsoft OneNote Desktop exposing the OneNote.Application COM API
              - .NET 8 runtime, or publish as a self-contained executable

            Commands:
              onenotemd import <input.md|directory> [output.one]
              onenotemd export <input.one> [output.md|directory]
              onenotemd convert <input> [output]
              onenotemd selftest

            Notes:
              - A .one file is a OneNote section and may contain several pages.
              - Exporting a multi-page section to a path ending in .md is rejected.
                Use a directory instead.
              - Importing a directory creates one OneNote page per .md file.
              - Existing output .one files are never overwritten.
            """);
    }
}
