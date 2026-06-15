# File Optimizer for AI

Convert documents, spreadsheets, presentations, images, audio, and other
supported files into clean Markdown, or learn repeatable Excel transformations
from example workbooks, through one Windows desktop application.

File Optimizer for AI is a graphical interface built around
[Microsoft MarkItDown](https://github.com/microsoft/markitdown). It is designed
for people who want AI-friendly text without using a command line or manually
copying content from complex files.

See [Technical Documentation](docs/TECHNICAL.md) for the internal architecture,
algorithms, threading model, packaging process, and extension guide.

Markdown preserves useful structure such as headings, lists, tables, and links
while keeping formatting lightweight. This makes the converted content easier
to review, search, and provide to tools such as ChatGPT, Claude, Gemini, and
other large language models.

## Download

Download the newest Windows installer from the
[[GitHub Releases page](../../releases/latest).](https://github.com/tareqRahman23/AI_File_Optimizer/releases/tag/AiFileOptimizer)

Run the downloaded setup file and follow the installation wizard. The installer
can create Start Menu and desktop shortcuts and includes the application
dependencies, so a separate Python installation is not required.

> Windows may display a SmartScreen warning for an unsigned installer. Review
> the publisher and file source before continuing.

## Features

- Simple native Windows desktop interface
- Drag-and-drop support for files and folders
- Select multiple files and convert them as one batch
- Add every supported file from a folder
- Parallel conversion using available CPU cores
- Per-file conversion status
- Estimated token reduction after conversion
- Custom output folder selection
- Automatic numbered filenames to prevent overwriting existing output
- Responsive background processing
- Subtle interface animations
- Local file conversion using Microsoft MarkItDown
- Separate example-based Excel formatting workspace
- Infer verified Excel transformations without an LLM or internet connection
- Save and reuse Excel transformation templates
- Preserve selected workbook styling, formulas, filters, sorting, and grouping

## Example-Based Excel Formatter

The **Excel Formatter** tab learns a reusable transformation from:

1. An example input workbook.
2. The correctly formatted output workbook.

The app compares the pair and keeps only rules that reproduce the expected
output exactly. You can review the detected actions, save them as a JSON
template, apply them to another workbook, or export a reproduced output.

Supported transformations include column selection, removal, reordering and
renaming; simple filters; grouping and aggregation; sorting; arithmetic
calculated columns; multi-sheet processing; and selected output styling such
as header formatting, column widths, number formats, freeze panes, and filters.

The formatter supports `.xlsx` and `.xlsm` input workbooks. It behaves
conservatively and reports an error instead of guessing when an example cannot
be represented safely.

## Supported Formats

The application currently accepts:

| Category | Formats |
| --- | --- |
| Documents | PDF, DOCX, RTF, EPUB |
| Spreadsheets and data | XLS, XLSX, CSV, JSON, XML |
| Presentations | PPTX |
| Web and text | HTML, HTM, TXT, Markdown |
| Images | PNG, JPG, JPEG, GIF, WEBP |
| Audio | MP3, WAV, M4A |
| Archives | ZIP |

Conversion quality depends on the source file and the converter available in
MarkItDown. Some formats may require additional system tools, such as FFmpeg
for certain audio files.

## Why Markdown?

Markdown is close to plain text but can still represent important document
structure. It is readable by people, easy to edit, and widely understood by
modern AI systems.

Compared with uploading or pasting raw document content, Markdown can:

- Remove unnecessary visual formatting
- Preserve headings, lists, tables, and links
- Make extracted content easier to inspect
- Reduce the approximate number of tokens needed to represent a document
- Work well in prompts, knowledge bases, search indexes, and text pipelines

The token reduction shown by the application is an estimate based on a general
character and file-size approximation. It is not an exact count for any
specific model or tokenizer.

## Using the Application

1. Open **File Optimizer for AI**.
2. Drag files into the drop area, or select **Add files**.
3. Use **Add folder** to queue supported files from one folder.
4. Review the output folder or choose a different location.
5. Select **Convert to Markdown**.
6. Review the status of each file and open the output folder when finished.

Converted files use the `.md` extension. If a filename already exists, the app
adds a number such as `_2` instead of replacing the existing file.

## Windows Permissions

The default output location is inside the current user's Documents folder and
normally does not require administrator access.

If you choose a protected location such as `C:\Program Files`, Windows may
prevent the app from creating files there. In that situation:

1. Close File Optimizer for AI.
2. Right-click its shortcut.
3. Select **Run as administrator**.

Alternatively, choose an output folder inside Documents, Desktop, or another
location where your Windows account has write permission.

## Limitations

- Complex layouts may not match the original document exactly.
- Scanned or image-only PDFs may require OCR support.
- Password-protected, damaged, or unsupported files may fail to convert.
- Audio conversion may require FFmpeg or other optional dependencies.
- Large batches can use significant memory and CPU resources.
- Estimated token reduction is approximate.
- The output is optimized for text processing, not high-fidelity visual
  reproduction.

## Run from Source

### Prerequisites

- Windows 10 or Windows 11
- Python 3.10 through 3.13
- An internet connection for the initial dependency installation

Python 3.13 is recommended. During Python installation, enable
**Add Python to PATH**.

### Automatic Setup

Clone or download the repository, then run:

```bat
install.bat
```

This creates a private `.tkvenv` environment and installs PySide6,
Microsoft MarkItDown, and the supported conversion dependencies.

After installation, start the application with:

```bat
run.bat
```

### Manual Setup

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python app.py
```

## Development

Run the test suite:

```powershell
python -m unittest discover -s tests
```

Run the application:

```powershell
python app.py
```

The main interface is located in `markitdown_gui/app.py`, and conversion logic
is located in `markitdown_gui/converter.py`.

## Building the Windows Installer

The Windows release is packaged with PyInstaller and Inno Setup.

Install the build dependencies:

```powershell
python -m pip install -r requirements-build.txt
```

Build the application and installer:

```powershell
.\scripts\build_windows_installer.ps1
```

The generated setup file is written to the `installer` directory.

Build artifacts and installers should only be regenerated when preparing a new
release.

## Security

File Optimizer for AI processes files with the permissions of the current
Windows user. Only convert files you trust.

The application uses MarkItDown's local-file conversion method and disables
third-party plugins. It does not intentionally upload local files through its
own interface. However, optional converters and external tools supplied by
dependencies may have their own behavior and requirements.

Avoid running the app as administrator unless access to a protected output
folder is genuinely required.

## Troubleshooting

### The app cannot write to the output folder

Choose a folder inside Documents or run the application as administrator if
you intentionally need to write to a protected Windows directory.

### A file fails to convert

Confirm that the file opens normally in its original application and is not
password-protected or damaged. Some formats require optional tools or contain
layouts that MarkItDown cannot fully extract.

### The source version does not start

Run `install.bat` again and review the displayed error. Confirm that Python
3.10 through 3.13 is installed and available on `PATH`.

### Audio conversion fails

Install FFmpeg and ensure it is available on the system `PATH`, then restart
the application.

## Credits

File Optimizer for AI was created by **Tareq Rahman Pantho**.

The conversion engine is
[Microsoft MarkItDown](https://github.com/microsoft/markitdown), an independent
open-source project maintained by Microsoft and its contributors.

This application is not an official Microsoft product and does not imply
Microsoft sponsorship or endorsement.

## Contributing

Bug reports, feature requests, and pull requests are welcome. When reporting a
conversion issue, include:

- The source file type
- The application version
- The displayed error message
- Whether the installed release or source version was used

Do not attach private or sensitive documents to public issues. Create a small,
sanitized sample file when reproduction material is needed.

## License and Third-Party Software

The application is provided as-is. Before distributing or accepting outside
contributions, add a repository license that clearly states how this project's
source code may be used.

Microsoft MarkItDown and other bundled dependencies are distributed under
their respective licenses. Refer to each dependency's project and packaged
license information for details.
