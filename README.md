# DOCX Proposal Generator

DOCX Proposal Generator is a Windows GUI app for generating polished proposal
documents from Markdown while preserving a real Word front matter template.

The app combines:

1. A front matter `.docx` or `.dotx` template
2. A Pandoc reference `.docx` for body styling
3. A source Markdown file
4. GUI metadata fields
5. A final `.docx` output path

The final output is always saved as `.docx`.

## Features

- Windows `tkinter` GUI
- Front matter templates in `.docx` or `.dotx` format
- Metadata fields for `Client Name`, `Document Name`, and `Author`
- Placeholder replacement for:
  - `{{CLIENT_NAME}}`
  - `{{DOCUMENT_NAME}}`
  - `{{AUTHOR}}`
  - `{{BODY_CONTENT}}` as the body insertion marker only
- OOXML-level placeholder patching before and after Word composition
- Markdown body conversion through bundled Pandoc
- Existing Word table of contents reuse
- Duplicate TOC cleanup in the Word automation path
- Word COM automation for field, TOC, header, footer, and page number updates
- Fallback DOCX generation when Microsoft Word is unavailable
- PyInstaller build script

## Requirements

- Windows
- Python 3.10 or newer
- Local project virtual environment in `.venv`

Python dependencies are listed in `requirements.txt`:

```text
pypandoc_binary
python-docx
pywin32
```

Pandoc is provided by `pypandoc_binary`, so Pandoc does not need to be installed
in the system `PATH`.

Microsoft Word is optional but recommended. If Word is installed, the app uses
Word COM automation to produce the best final document. If Word is not
available, the app still generates a DOCX where possible and warns that fields
and TOC page numbers may need to be updated manually in Word.

## Setup

Create and install the local virtual environment:

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

Verify the bundled Pandoc binary:

```powershell
.\.venv\Scripts\python.exe -c "import pypandoc; print(pypandoc.get_pandoc_path())"
```

## Run

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
page, logos, tables, watermark, headers, footers, page numbering, document
control page, and table of contents you want to preserve.

Supported placeholders:

```text
{{CLIENT_NAME}}
{{DOCUMENT_NAME}}
{{AUTHOR}}
{{BODY_CONTENT}}
```

`{{CLIENT_NAME}}`, `{{DOCUMENT_NAME}}`, and `{{AUTHOR}}` are replaced from the
GUI metadata fields.

`{{BODY_CONTENT}}` is not replaced with metadata. It is used only as the marker
where the converted Markdown body is inserted.

Placeholder replacement is performed at OOXML package level and through Word COM
where available, so it can handle normal body text, tables, headers, footers,
text boxes, shapes, and grouped shapes.

If the front matter file is `.dotx`, Microsoft Word creates a new document with:

```text
Word.Documents.Add(Template=template_path)
```

If it is `.docx`, Word opens it with:

```text
Word.Documents.Open(template_path)
```

Do not use `.dotx` as the Pandoc reference document. The Pandoc reference input
must remain `.docx` for compatibility.

## Table Of Contents

The front matter template should contain the Word TOC field if a TOC is needed.
The app prefers the existing TOC from the template and does not create another
one during Word composition.

In the Word automation path:

- the existing TOC is updated after the body is inserted
- if multiple TOCs exist, all TOCs except the first are deleted
- all fields, headers, footers, and page numbers are updated before saving

Pandoc is not called with `--toc`.

## Generate A Proposal

1. Click `Browse Front Matter` and select the `.docx` or `.dotx` front matter template.
2. Click `Browse Reference DOCX` and select the Pandoc reference `.docx`.
3. Click `Browse Source MD` and select the Markdown body file.
4. Click `Save Output As` and choose the final `.docx` path.
5. Enter `Client Name`, `Document Name`, and `Author`.
6. Click `Start`.

The body conversion command is:

```bash
pandoc source.md --reference-doc=reference.docx -o body.docx
```

The app then:

1. patches metadata placeholders in a temporary copy of the front matter template
2. converts Markdown to a temporary `body.docx`
3. creates or opens the front matter document
4. inserts `body.docx` at `{{BODY_CONTENT}}`
5. updates the existing TOC and Word fields when Word is available
6. saves the final `.docx`
7. runs a final OOXML safety replacement pass on the saved output

## Build

Run:

```bat
build.bat
```

This installs PyInstaller into the local `.venv` and creates:

```text
dist\DOCXProposalGenerator.exe
```

## Notes And Limitations

Do not try to extract "the first two pages" from another DOCX. DOCX pagination is
dynamic, so front matter should be provided as its own `.docx` or `.dotx`
template.

The Word COM path gives the best preservation of complex cover pages, headers,
footers, fields, TOCs, page numbering, tables, and Word-specific layout.

The fallback path uses `python-docx` and direct OOXML patching. It can generate a
usable DOCX without Microsoft Word, but field updates, TOC page numbers, and some
complex document relationships may require opening the document in Word and
updating fields.

The Pandoc reference `.docx` controls body styles, fonts, margins, and general
body formatting. It does not guarantee perfect replication of complex custom
layouts unless the Markdown structure maps cleanly to the reference document
styles.
