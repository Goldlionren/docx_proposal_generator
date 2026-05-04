import os
import shutil
import subprocess
import sys
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, ttk


APP_NAME = "DOCX Proposal Generator"
APP_DIR = Path(__file__).resolve().parent


class DOCXProposalGenerator(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title(APP_NAME)
        self.geometry("850x520")
        self.minsize(720, 440)

        self.reference_docx = tk.StringVar()
        self.source_markdown = tk.StringVar()
        self.output_docx = tk.StringVar()

        self._build_ui()

    def _build_ui(self):
        root = ttk.Frame(self, padding=16)
        root.pack(fill=tk.BOTH, expand=True)

        title = ttk.Label(root, text=APP_NAME, font=("Segoe UI", 16, "bold"))
        title.grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 16))

        root.columnconfigure(1, weight=1)
        root.rowconfigure(5, weight=1)

        self._add_file_row(
            root,
            row=1,
            label="Reference DOCX:",
            variable=self.reference_docx,
            button_text="Browse Reference DOCX",
            command=self.browse_reference_docx,
        )
        self._add_file_row(
            root,
            row=2,
            label="Source Markdown:",
            variable=self.source_markdown,
            button_text="Browse Source MD",
            command=self.browse_source_markdown,
        )
        self._add_file_row(
            root,
            row=3,
            label="Output DOCX:",
            variable=self.output_docx,
            button_text="Save Output As",
            command=self.browse_output_docx,
        )

        start_button = ttk.Button(root, text="Start", command=self.start_conversion)
        start_button.grid(row=4, column=0, sticky="w", pady=(12, 12))

        status_label = ttk.Label(root, text="Status:")
        status_label.grid(row=5, column=0, sticky="nw", pady=(0, 4))

        self.status_text = scrolledtext.ScrolledText(root, height=12, wrap=tk.WORD)
        self.status_text.grid(row=5, column=1, columnspan=2, sticky="nsew")
        self.status_text.configure(state=tk.DISABLED)

    def _add_file_row(self, parent, row, label, variable, button_text, command):
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", pady=6)

        entry = ttk.Entry(parent, textvariable=variable, state="readonly")
        entry.grid(row=row, column=1, sticky="ew", padx=(8, 8), pady=6)

        button = ttk.Button(parent, text=button_text, command=command)
        button.grid(row=row, column=2, sticky="ew", pady=6)

    def browse_reference_docx(self):
        path = filedialog.askopenfilename(
            title="Select Reference DOCX",
            filetypes=[("Word Documents", "*.docx")],
        )
        if path:
            self.reference_docx.set(path)
            self.log(f"Selected reference file: {path}")

    def browse_source_markdown(self):
        path = filedialog.askopenfilename(
            title="Select Source Markdown",
            filetypes=[
                ("Markdown Files", "*.md *.markdown"),
                ("All Files", "*.*"),
            ],
        )
        if path:
            self.source_markdown.set(path)
            self.log(f"Selected source file: {path}")

    def browse_output_docx(self):
        path = filedialog.asksaveasfilename(
            title="Save Output DOCX As",
            defaultextension=".docx",
            filetypes=[("Word Documents", "*.docx")],
        )
        if path:
            self.output_docx.set(path)
            self.log(f"Selected output file: {path}")

    def start_conversion(self):
        reference_docx = self.reference_docx.get().strip()
        source_markdown = self.source_markdown.get().strip()
        output_docx = self.output_docx.get().strip()

        if not self._is_file(reference_docx):
            self.show_error("Please select a valid reference DOCX file.")
            return

        if not self._is_file(source_markdown):
            self.show_error("Please select a valid source Markdown file.")
            return

        if not output_docx:
            self.show_error("Please select an output DOCX path.")
            return

        if not self._has_docx_extension(output_docx):
            self.show_error("Please select an output DOCX path ending in .docx.")
            return

        output_parent = Path(output_docx).expanduser().parent
        if output_parent and not output_parent.exists():
            self.show_error(f"Output folder does not exist:\n{output_parent}")
            return

        if os.path.exists(output_docx):
            should_overwrite = messagebox.askyesno(
                "Confirm Overwrite",
                "The output file already exists. Do you want to overwrite it?",
            )
            if not should_overwrite:
                self.log("Conversion cancelled. Output file already exists.")
                return

        pandoc_path = self._find_pandoc()
        if not pandoc_path:
            self.show_error(
                "Pandoc is not installed in the local Python environment.\n"
                "Please run: .\\.venv\\Scripts\\python.exe -m pip install -r requirements.txt"
            )
            return

        if not self._check_pandoc_version(pandoc_path):
            return

        pandoc_command = [
            pandoc_path,
            source_markdown,
            "--reference-doc",
            reference_docx,
            "-o",
            output_docx,
        ]

        self.log("Running Pandoc command:")
        self.log(self._format_command(pandoc_command))

        try:
            result = subprocess.run(
                pandoc_command,
                capture_output=True,
                text=True,
                check=False,
            )
        except OSError as exc:
            self.show_error(f"Failed to run Pandoc:\n{exc}")
            return

        if result.stdout:
            self.log("Pandoc stdout:")
            self.log(result.stdout.strip())

        if result.stderr:
            self.log("Pandoc stderr:")
            self.log(result.stderr.strip())

        if result.returncode == 0:
            self.log(f"Success. DOCX generated at: {output_docx}")
            messagebox.showinfo("Success", "DOCX generated successfully.")
        else:
            self.show_error(
                "Pandoc conversion failed.\n\n"
                f"Return code: {result.returncode}\n\n"
                f"{result.stderr.strip() or 'No error output was returned.'}"
            )

    def _check_pandoc_version(self, pandoc_path):
        try:
            result = subprocess.run(
                [pandoc_path, "--version"],
                capture_output=True,
                text=True,
                check=False,
            )
        except OSError:
            self.show_error(
                "Pandoc is not installed in the local Python environment.\n"
                "Please run: .\\.venv\\Scripts\\python.exe -m pip install -r requirements.txt"
            )
            return False

        if result.returncode != 0:
            self.show_error(
                "Pandoc is not installed in the local Python environment.\n"
                "Please run: .\\.venv\\Scripts\\python.exe -m pip install -r requirements.txt"
            )
            return False

        version_line = result.stdout.splitlines()[0] if result.stdout else "pandoc"
        self.log(f"Pandoc available: {version_line}")
        self.log(f"Pandoc path: {pandoc_path}")
        return True

    def _find_pandoc(self):
        candidate_paths = [
            APP_DIR / ".venv" / "Scripts" / "pandoc.exe",
            Path(sys.prefix) / "Scripts" / "pandoc.exe",
            APP_DIR / "pandoc.exe",
            APP_DIR / "tools" / "pandoc.exe",
            APP_DIR / "tools" / "pandoc" / "pandoc.exe",
        ]

        for candidate_path in candidate_paths:
            if candidate_path.is_file():
                return str(candidate_path)

        try:
            import pypandoc

            pandoc_path = pypandoc.get_pandoc_path()
            if pandoc_path:
                pypandoc_candidate = Path(pandoc_path)
                if pypandoc_candidate.is_file():
                    return str(pypandoc_candidate)

                exe_candidate = pypandoc_candidate.with_suffix(".exe")
                if exe_candidate.is_file():
                    return str(exe_candidate)
        except (ImportError, OSError):
            pass

        return shutil.which("pandoc")

    def _is_file(self, path):
        return bool(path) and os.path.isfile(path)

    def _has_docx_extension(self, path):
        return Path(path).suffix.lower() == ".docx"

    def _format_command(self, command):
        return subprocess.list2cmdline(command)

    def log(self, message):
        self.status_text.configure(state=tk.NORMAL)
        self.status_text.insert(tk.END, f"{message}\n")
        self.status_text.see(tk.END)
        self.status_text.configure(state=tk.DISABLED)

    def show_error(self, message):
        self.log(f"Error: {message}")
        messagebox.showerror("Error", message)


def main():
    app = DOCXProposalGenerator()
    app.mainloop()


if __name__ == "__main__":
    main()
