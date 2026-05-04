# DOCX Proposal Generator

DOCX Proposal Generator is a simple Windows GUI app that creates a new `.docx`
proposal document from:

1. A reference/template `.docx`
2. A source Markdown file
3. An output `.docx` save path

The app uses Pandoc as the conversion engine:

```bash
pandoc source.md --reference-doc=reference.docx -o output.docx
```

## Requirements

- Windows
- Python 3.10 or newer
- Local project virtual environment in `.venv`

The GUI uses `tkinter`, which is included with standard Python installations.
Pandoc is provided by the Python package `pypandoc_binary`, so the app does not
require Pandoc to be installed in the system `PATH`.

## Set Up The Local Python Environment

From this folder, create and install the local virtual environment:

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

Verify the local Pandoc binary:

```powershell
.\.venv\Scripts\python.exe -c "import pypandoc; print(pypandoc.get_pandoc_path())"
```

## Run The App

From this folder, run:

```powershell
.\.venv\Scripts\python.exe main.py
```

## Generate An Output DOCX

1. Click `Browse Reference DOCX` and select the branded Word reference document.
2. Click `Browse Source MD` and select the Markdown content file.
3. Click `Save Output As` and choose where the generated `.docx` should be saved.
4. Click `Start`.

The status box shows selected files, the Pandoc command, success messages, and
error output if conversion fails.

If the selected output file already exists, the app asks before overwriting it.

## Build A Windows EXE

Run:

```bat
build.bat
```

This installs PyInstaller and creates a single-file Windows executable named
`DOCXProposalGenerator.exe` under the generated `dist` folder.

The build script uses the local `.venv` Python environment. Pandoc is resolved
from the local Python package first, then from project-local executable paths,
then from system `PATH` only as a final fallback.

## Known Limitations

The reference DOCX controls Word styles, fonts, headers, footers, margins, and
general formatting. It does not guarantee perfect replication of complex cover
pages or custom layouts unless the Markdown structure matches the reference
document styles.

V1 does not:

- Use Microsoft Word automation
- Require Microsoft Office
- Modify DOCX XML directly
- Replace Word placeholders
- Export to PDF
- Analyze or map Word styles in the GUI
