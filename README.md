# CSharp-OneNote-To-Markdown

A small Windows CLI that converts **Markdown ↔ Microsoft OneNote `.one` sections** without any third-party runtime dependency.

It intentionally does **not** parse or generate Microsoft's proprietary `.one` binary format itself. Instead, it asks the installed Microsoft OneNote Desktop application to open/create the section through the official `OneNote.Application` COM automation interface.

## Why this project exists

Many OneNote/Markdown converters depend on third-party add-ins, Pandoc, Python packages, Node packages, or unofficial `.one` parsers. That is not always acceptable on locked-down workstations.

This project only needs:

- Windows
- Microsoft OneNote Desktop
- .NET 8 (or a self-contained published executable)
- this repository's own source code

There are **no `PackageReference` entries** and no external runtime libraries.

## Commands

```powershell
onenotemd import notes.md notes.one
onenotemd import .\notes-directory\ notes.one

onenotemd export notes.one .\markdown-output\
onenotemd export single-page.one page.md

onenotemd convert notes.md
onenotemd convert notes.one

onenotemd selftest
```

### Important `.one` detail

A `.one` file is a **OneNote section**, not a single page. A section can contain many pages.

Therefore:

- importing one `.md` creates one page in a new `.one` section;
- importing a directory creates one page per Markdown file;
- exporting a multi-page `.one` creates one Markdown file per page;
- exporting to a single `output.md` is accepted only when the `.one` section contains exactly one page.

## Build

```powershell
dotnet build -c Release
```

Run:

```powershell
dotnet run -- convert example.md example.one
```

Publish a single executable that does not require a preinstalled .NET runtime:

```powershell
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true
```

The resulting executable still requires **Microsoft OneNote Desktop**, because OneNote itself is deliberately used as the `.one` reader/writer.

## Supported Markdown in the first version

The built-in converter handles the common note-taking subset:

- headings (`#` to `######`)
- paragraphs
- bold
- italic
- strikethrough
- inline code
- fenced code blocks
- links
- unordered lists
- ordered lists
- task-list markers (`[ ]`, `[x]`)
- simple pipe tables
- quotes (rendered as emphasized paragraphs in OneNote)
- standalone local images

No Markdown parsing library is used.

### Images

A local image written on its own line is embedded into OneNote:

```markdown
![diagram](images/diagram.png)
```

Supported embedded image formats are PNG, JPEG, GIF, BMP, and TIFF.

HTTP/HTTPS images are left as links instead of being downloaded automatically. This keeps conversion deterministic and avoids unexpected network access.

When exporting from OneNote, embedded image data exposed by the COM API is written beside the Markdown file in a `.assets` directory.

## OneNote → Markdown behavior

The exporter retrieves the page hierarchy and page XML from OneNote and flattens OneNote's free-form canvas into Markdown.

Outlines are sorted by their `(y, x)` position before conversion. This is necessarily lossy because Markdown has a linear document model while OneNote allows independently positioned content containers.

The exporter currently preserves the useful semantic subset:

- page title
- text and common inline formatting
- lists
- simple tables
- images

Ink, audio/video, advanced tags, equations, handwriting layout, file attachments, and exact canvas positioning are not represented faithfully in Markdown.

## Markdown → OneNote behavior

The importer:

1. creates a `.one` section with `OpenHierarchy`;
2. creates one OneNote page per Markdown file with `CreateNewPage`;
3. converts Markdown to safe HTML using code in this repository;
4. inserts that HTML into OneNote through OneNote XML;
5. embeds supported local images as base64 OneNote image objects.

The first top-level `# Heading` in a Markdown file is used as the OneNote page title and removed from the body. If there is no H1, the file name becomes the page title.

For safety, the CLI refuses to overwrite an existing `.one` output file.

## Why late-bound COM instead of `Microsoft.Office.Interop.OneNote`

Using the interop assembly would add another binary/reference that has to be approved and deployed.

This project resolves the registered COM class at runtime:

```csharp
Type.GetTypeFromProgID("OneNote.Application")
```

and invokes the documented OneNote automation methods through reflection. The implementation therefore compiles without the Office interop NuGet package or a copied interop DLL.

## Security / deployment characteristics

- no third-party libraries;
- no network access is needed for normal local conversion;
- remote Markdown images are not downloaded;
- existing `.one` files are not overwritten;
- conversion runs locally through the installed OneNote desktop client.

Set `ONENOTEMD_DEBUG=1` to print full exception details.

## Known limitations

This is intentionally a pragmatic converter rather than a bit-perfect OneNote clone.

Markdown and OneNote have different document models, so round-tripping cannot preserve every feature. In particular:

- OneNote canvas positions become a linear Markdown order;
- complex nested list styling may simplify;
- advanced OneNote tags and ink are not preserved;
- Markdown images are embedded only when they are standalone local-image lines;
- complex HTML generated internally by older OneNote versions may be simplified during export.

The code is structured so the Markdown and OneNote-XML converters can be improved independently while keeping the zero-third-party-dependency constraint.
