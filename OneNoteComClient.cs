using System.Reflection;
using System.Runtime.InteropServices;

namespace OneNoteMarkdown;

internal sealed class OneNoteComClient : IDisposable
{
    private const int CftNone = 0;
    private const int CftSection = 3;
    private const int HsPages = 4;
    private const int NpsBlankPageWithTitle = 1;
    private const int PiAll = 7;
    private const int Xs2013 = 2;

    private readonly object _application;
    private readonly Type _applicationType;
    private bool _disposed;

    public OneNoteComClient()
    {
        if (!OperatingSystem.IsWindows())
        {
            throw new PlatformNotSupportedException("The OneNote COM API is only available on Windows.");
        }

        _applicationType = Type.GetTypeFromProgID("OneNote.Application")
            ?? throw new InvalidOperationException(
                "Microsoft OneNote Desktop is not installed, or the OneNote.Application COM class is not registered.");

        _application = Activator.CreateInstance(_applicationType)
            ?? throw new InvalidOperationException("Could not create OneNote.Application.");
    }

    public string OpenSection(string path, bool create)
    {
        object?[] args = [Path.GetFullPath(path), string.Empty, null, create ? CftSection : CftNone];
        Invoke("OpenHierarchy", args);
        return RequireString(args[2], "OneNote did not return a section ID.");
    }

    public string GetPageHierarchy(string sectionId)
    {
        object?[] args = [sectionId, HsPages, null, Xs2013];
        Invoke("GetHierarchy", args);
        return RequireString(args[2], "OneNote did not return hierarchy XML.");
    }

    public string CreatePage(string sectionId)
    {
        object?[] args = [sectionId, null, NpsBlankPageWithTitle];
        Invoke("CreateNewPage", args);
        return RequireString(args[1], "OneNote did not return a page ID.");
    }

    public string GetPageContent(string pageId)
    {
        object?[] args = [pageId, null, PiAll, Xs2013];
        Invoke("GetPageContent", args);
        return RequireString(args[1], "OneNote did not return page XML.");
    }

    public string? GetBinaryPageContent(string pageId, string callbackId)
    {
        object?[] args = [pageId, callbackId, null];
        Invoke("GetBinaryPageContent", args);
        return args[2] as string;
    }

    public void UpdatePageContent(string pageXml)
    {
        object?[] args = [pageXml, DateTime.MinValue, Xs2013, false];
        Invoke("UpdatePageContent", args);
    }

    private void Invoke(string method, object?[] args)
    {
        try
        {
            _applicationType.InvokeMember(
                method,
                BindingFlags.InvokeMethod | BindingFlags.Instance | BindingFlags.Public,
                binder: null,
                target: _application,
                args: args);
        }
        catch (TargetInvocationException ex) when (ex.InnerException is not null)
        {
            throw new InvalidOperationException(
                $"OneNote COM call {method} failed: {ex.InnerException.Message}",
                ex.InnerException);
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException(
                $"OneNote COM call {method} failed (0x{ex.HResult:X8}): {ex.Message}",
                ex);
        }
    }

    private static string RequireString(object? value, string message) =>
        value as string is { Length: > 0 } text
            ? text
            : throw new InvalidOperationException(message);

    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        if (Marshal.IsComObject(_application))
        {
            Marshal.FinalReleaseComObject(_application);
        }

        _disposed = true;
        GC.SuppressFinalize(this);
    }
}
