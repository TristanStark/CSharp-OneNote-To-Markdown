namespace OneNoteMarkdown;

internal static class Program
{
    [STAThread]
    private static int Main(string[] args)
    {
        try
        {
            return Cli.Run(args);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"error: {ex.Message}");
            if (Environment.GetEnvironmentVariable("ONENOTEMD_DEBUG") == "1")
            {
                Console.Error.WriteLine(ex);
            }

            return 1;
        }
    }
}
