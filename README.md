# DOCX Proposal Generator

DOCX Proposal Generator is a Windows GUI app for producing proposal `.docx`
files from:

1. A front matter Word template or document
2. A Pandoc reference Word document
3. A source Markdown file
4. Client/document metadata entered in the GUI
5. An output `.docx` save path

The Markdown body is converted with Pandoc, then inserted into the selected front
matter template at `{{BODY_CONTENT}}`.

## Requirements

- Windows
- Python 3.10 or newer
- Local project virtual environment in `.venv`

The app uses `tkinter` for the GUI. Pandoc is provided by `pypandoc_binary`, so
Pandoc does not need to be installed in the system `PATH`.

Microsoft Word is optional but recommended. When Word is installed, the app uses
Word COM automation through `pywin32` to update fields, the table of contents,
headers, footers, and page numbers before saving. When Word is not installed,
the app still generates a DOCX where possible and warns that the TOC may need to
be updated manually in Word.

## Set Up The Local Python Environment

From this folder, create and install the local virtual environment:

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

Verify the bundled Pandoc binary:

```powershell
.\.venv\Scripts\python.exe -c "import pypandoc; print(pypandoc.get_pandoc_path())"
```

## Run The App

Use the helper script:

```bat
run.bat
```

Or run directly:

```powershell
.\.venv\Scripts\python.exe main.py
```

## Front Matter Template

Create a real Word `.docx` or `.dotx` front matter template containing the cover
page, logos, tables, watermark, headers, footers, and page numbering you want to
preserve.

The template may contain these placeholders:

```text
{{CLIENT_NAME}}
{{DOCUMENT_NAME}}
{{AUTHOR}}
{{BODY_CONTENT}}
```

Only those placeholders are replaced. Other table cells and template structures
are left unchanged. The body insertion point is `{{BODY_CONTENT}}`.

If the front matter file is `.dotx`, Microsoft Word creates a new document from
that template. If it is `.docx`, Word opens the document directly. The final
output is always saved as `.docx`.

When Microsoft Word is not available, `.dotx` templates are loaded through a
temporary `.docx` package copy so the fallback composer can still generate an
output document.

Do not try to extract "the first two pages" from another DOCX. DOCX pagination is
dynamic, so the front matter must be supplied as its own template.

## Generate An Output DOCX

1. Click `Browse Front Matter` and select the `.docx` or `.dotx` front matter template.
2. Click `Browse Reference DOCX` and select the Pandoc reference `.docx`.
3. Click `Browse Source MD` and select the Markdown body file.
4. Click `Save Output As` and choose where the generated `.docx` should be saved.
5. Enter `Client Name`, `Document Name`, and `Author`.
6. Click `Start`.

The app converts the Markdown body with:

```bash
pandoc source.md --reference-doc=reference.docx -o body.docx
```

It then inserts a Word table of contents before the body and inserts `body.docx`
at `{{BODY_CONTENT}}`.

If the front matter template already contains a Word table of contents field,
the app uses that existing TOC and does not create a duplicate. If multiple TOCs
exist in the template, the Word automation path keeps the first one and removes
the rest before updating fields.

Do not use `.dotx` as the Pandoc reference document. The Pandoc reference input
is intentionally kept as `.docx` for compatibility.

## Build A Windows EXE

Run:

```bat
build.bat
```

This installs PyInstaller into the local `.venv` and creates a single-file
Windows executable named `DOCXProposalGenerator.exe` under `dist`.

## Known Limitations

The Word COM path gives the best preservation of complex cover pages, headers,
footers, fields, and page numbering.

The fallback path uses `python-docx`. It can generate a usable DOCX without
Microsoft Word, but field updates, TOC page numbers, and some complex body
relationships may require opening the document in Word and updating fields.

The Pandoc reference DOCX controls Word styles, fonts, margins, and general body
formatting. It does not guarantee perfect replication of complex custom layouts
unless the Markdown structure maps cleanly to the reference document styles.
