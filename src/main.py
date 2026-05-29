#!/usr/bin/env python3
"""
File → Markdown Converter
Sin dependencias externas de DnD — compatible Linux y Windows.
Requiere: pip install markitdown[all]
"""

import os
import sys
import shutil
import subprocess
import queue
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from urllib.parse import urlparse, unquote
from urllib.request import url2pathname

# ── Dependencia principal ──────────────────────────────────────────────────────
try:
    from markitdown import MarkItDown
    HAS_MARKITDOWN = True
except ImportError:
    HAS_MARKITDOWN = False

# ── Constantes ────────────────────────────────────────────────────────────────
SUPPORTED_EXT = (
    ".pdf", ".docx", ".doc", ".pptx", ".ppt",
    ".xlsx", ".xls", ".html", ".htm",
    ".csv", ".txt", ".xml", ".json",
    ".epub", ".zip", ".md",
)

FILETYPES = [
    ("Archivos soportados",
     " ".join(f"*{e}" for e in SUPPORTED_EXT)),
    ("PDF",              "*.pdf"),
    ("Word",             "*.docx *.doc"),
    ("PowerPoint",       "*.pptx *.ppt"),
    ("Excel",            "*.xlsx *.xls"),
    ("HTML",             "*.html *.htm"),
    ("Texto / CSV / XML","*.txt *.csv *.xml *.json"),
    ("Todos los arch.",  "*.*"),
]

HINT = (
    "PDF · DOCX · PPTX · XLSX · HTML · CSV · TXT · XML · EPUB · ZIP\n"
    "Haz clic · pega la ruta abajo · o usa Ctrl+O"
)

COLOR_IDLE   = "#ddeeff"
COLOR_HOVER  = "#b8d8ff"


# ── Diálogos nativos del SO ───────────────────────────────────────────────────
# En Windows y macOS, filedialog ya usa el selector nativo del sistema.
# En Linux detectamos zenity (GNOME/GTK) o kdialog (KDE) y los preferimos;
# si ninguno está instalado caemos al selector de Tkinter como último recurso.

def _detect_linux_dialog() -> str:
    """Devuelve 'zenity', 'kdialog', o 'tk' según lo disponible en el sistema."""
    if sys.platform != "linux":
        return "tk"
    if shutil.which("zenity"):
        return "zenity"
    if shutil.which("kdialog"):
        return "kdialog"
    return "tk"

_DIALOG = _detect_linux_dialog()

# Filtros para zenity:  "Descripción | *.ext1 *.ext2"
_Z_OPEN_FILTERS = [
    "Archivos soportados | " + " ".join(f"*{e}" for e in (
        ".pdf", ".docx", ".doc", ".pptx", ".ppt",
        ".xlsx", ".xls", ".html", ".htm",
        ".csv", ".txt", ".xml", ".json", ".epub", ".zip", ".md",
    )),
    "PDF | *.pdf",
    "Word | *.docx *.doc",
    "PowerPoint | *.pptx *.ppt",
    "Excel | *.xlsx *.xls",
    "HTML | *.html *.htm",
    "Texto / CSV / XML | *.txt *.csv *.xml *.json",
    "Todos los archivos | *",
]
_Z_SAVE_FILTERS = [
    "Markdown | *.md",
    "Texto | *.txt",
    "Todos los archivos | *",
]

# Filtros para kdialog: "*.ext1 *.ext2|Descripción\n*.ext3|Otra" (newlines reales)
_K_OPEN_FILTER = "\n".join([
    "*.pdf *.docx *.doc *.pptx *.ppt *.xlsx *.xls"
    " *.html *.htm *.csv *.txt *.xml *.json *.epub *.zip *.md"
    "|Archivos soportados",
    "*.pdf|PDF",
    "*.docx *.doc|Word",
    "*.pptx *.ppt|PowerPoint",
    "*.xlsx *.xls|Excel",
    "*.html *.htm|HTML",
    "*.txt *.csv *.xml *.json|Texto / CSV / XML",
    "*|Todos los archivos",
])
_K_SAVE_FILTER = "*.md|Markdown\n*.txt|Texto\n*|Todos los archivos"


def dialog_open() -> str | None:
    """Abre el selector nativo para elegir un archivo. Devuelve la ruta o None."""
    if _DIALOG == "zenity":
        cmd = ["zenity", "--file-selection", "--title=Abrir archivo"]
        for f in _Z_OPEN_FILTERS:
            cmd.append(f"--file-filter={f}")
        try:
            r = subprocess.run(cmd, capture_output=True, text=True)
            path = r.stdout.strip()
            return path or None
        except Exception:
            pass  # cae al fallback

    if _DIALOG == "kdialog":
        try:
            r = subprocess.run(
                ["kdialog", "--getopenfilename",
                 os.path.expanduser("~"), _K_OPEN_FILTER],
                capture_output=True, text=True,
            )
            path = r.stdout.strip()
            return path or None
        except Exception:
            pass

    # Fallback: Tkinter (no nativo en Linux, pero funcional)
    return filedialog.askopenfilename(filetypes=FILETYPES) or None


def dialog_save(initial_name: str = "output.md") -> str | None:
    """Abre el selector nativo para guardar. Devuelve la ruta elegida o None."""
    initial_path = os.path.join(os.path.expanduser("~"), initial_name)

    if _DIALOG == "zenity":
        cmd = [
            "zenity", "--file-selection", "--save",
            "--confirm-overwrite",
            "--title=Guardar como",
            f"--filename={initial_path}",
        ]
        for f in _Z_SAVE_FILTERS:
            cmd.append(f"--file-filter={f}")
        try:
            r = subprocess.run(cmd, capture_output=True, text=True)
            path = r.stdout.strip()
            if path:
                # Añade .md si el usuario no escribió extensión
                if not os.path.splitext(path)[1]:
                    path += ".md"
                return path
        except Exception:
            pass

    if _DIALOG == "kdialog":
        try:
            r = subprocess.run(
                ["kdialog", "--getsavefilename", initial_path, _K_SAVE_FILTER],
                capture_output=True, text=True,
            )
            path = r.stdout.strip()
            return path or None
        except Exception:
            pass

    # Fallback: Tkinter
    return filedialog.asksaveasfilename(
        defaultextension=".md",
        initialfile=initial_name,
        filetypes=[("Markdown", "*.md"), ("Texto", "*.txt"), ("Todos", "*.*")],
    ) or None


# ── Aplicación ────────────────────────────────────────────────────────────────
class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("File → Markdown Converter")
        self.root.geometry("860x680")
        self.root.minsize(620, 480)

        self.markdown_content = ""
        self.source_filename  = ""
        self._q: queue.Queue  = queue.Queue()

        self._build_ui()
        self._check_deps()
        self._poll()

        # Atajos globales
        self.root.bind("<Control-o>", self._browse)
        self.root.bind("<Control-O>", self._browse)

    # ── UI ────────────────────────────────────────────────────────────────────

    def _build_ui(self):
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        main = ttk.Frame(self.root, padding=14)
        main.grid(sticky="nsew")
        main.columnconfigure(0, weight=1)
        main.rowconfigure(4, weight=1)   # text area expands

        # Título
        ttk.Label(
            main,
            text="File → Markdown Converter",
            font=("TkDefaultFont", 15, "bold"),
        ).grid(row=0, column=0, sticky="w", pady=(0, 10))

        # ── Zona de apertura (click) ──────────────────────────────────────────
        self.dz = tk.Frame(
            main, bg=COLOR_IDLE, relief="ridge", bd=2,
            height=95, cursor="hand2",
        )
        self.dz.grid(row=1, column=0, sticky="ew", pady=(0, 8))
        self.dz.grid_propagate(False)
        self.dz.columnconfigure(0, weight=1)
        self.dz.rowconfigure(0, weight=1)

        self._lbl_hint = tk.Label(
            self.dz, text="📂  Haz clic aquí para abrir un archivo",
            bg=COLOR_IDLE, font=("TkDefaultFont", 12), justify="center",
        )
        self._lbl_hint.grid(row=0, column=0, pady=(12, 2))

        self._lbl_sub = tk.Label(
            self.dz, text=HINT,
            bg=COLOR_IDLE, font=("TkDefaultFont", 8), fg="#555", justify="center",
        )
        self._lbl_sub.grid(row=1, column=0, pady=(0, 8))

        for w in (self.dz, self._lbl_hint, self._lbl_sub):
            w.bind("<Button-1>", self._browse)
            w.bind("<Enter>",    lambda e: self._dz_color(COLOR_HOVER))
            w.bind("<Leave>",    lambda e: self._dz_color(COLOR_IDLE))

        # ── Campo de ruta / URI ───────────────────────────────────────────────
        path_row = ttk.Frame(main)
        path_row.grid(row=2, column=0, sticky="ew", pady=(0, 6))
        path_row.columnconfigure(1, weight=1)

        ttk.Label(path_row, text="Ruta:").grid(row=0, column=0, padx=(0, 6))

        self._path_var = tk.StringVar()
        self._path_entry = ttk.Entry(path_row, textvariable=self._path_var)
        self._path_entry.grid(row=0, column=1, sticky="ew")
        self._path_entry.bind("<Return>",  self._entry_go)
        self._path_entry.bind("<<Paste>>", self._entry_paste)

        ttk.Button(
            path_row, text="Convertir →", command=self._entry_go,
        ).grid(row=0, column=2, padx=(6, 0))

        ttk.Label(
            main,
            text='Tip: pega una ruta o un "file:// URI" y pulsa Enter  ·  Ctrl+O para explorar',
            foreground="#888", font=("TkDefaultFont", 8),
        ).grid(row=3, column=0, sticky="w", pady=(0, 2))

        # ── Barra de estado ───────────────────────────────────────────────────
        self._status = tk.StringVar(value="Listo.")
        ttk.Label(
            main, textvariable=self._status,
            foreground="#444", font=("TkDefaultFont", 9),
        ).grid(row=3, column=0, sticky="e")

        # ── Área de texto ─────────────────────────────────────────────────────
        out = ttk.LabelFrame(main, text="Markdown generado", padding=4)
        out.grid(row=4, column=0, sticky="nsew")
        out.columnconfigure(0, weight=1)
        out.rowconfigure(0, weight=1)

        self._text = scrolledtext.ScrolledText(
            out, wrap=tk.WORD, font=("Courier", 10),
            state=tk.DISABLED, relief="flat", bd=0,
        )
        self._text.grid(sticky="nsew")

        # ── Botones ───────────────────────────────────────────────────────────
        bar = ttk.Frame(main)
        bar.grid(row=5, column=0, sticky="ew", pady=(8, 0))

        self._btn_save  = ttk.Button(bar, text="💾 Guardar .md",  command=self._save,  state=tk.DISABLED)
        self._btn_copy  = ttk.Button(bar, text="📋 Copiar",        command=self._copy,  state=tk.DISABLED)
        self._btn_clear = ttk.Button(bar, text="🗑  Limpiar",       command=self._clear, state=tk.DISABLED)

        self._btn_save .pack(side=tk.LEFT, padx=(0, 6))
        self._btn_copy .pack(side=tk.LEFT, padx=(0, 6))
        self._btn_clear.pack(side=tk.LEFT)

        # ── Barra de progreso ─────────────────────────────────────────────────
        self._progress = ttk.Progressbar(main, mode="indeterminate")
        self._progress.grid(row=6, column=0, sticky="ew", pady=(6, 0))

    # ── Colores drop zone ─────────────────────────────────────────────────────

    def _dz_color(self, color: str):
        for w in (self.dz, self._lbl_hint, self._lbl_sub):
            w.config(bg=color)

    # ── Deps ──────────────────────────────────────────────────────────────────

    def _check_deps(self):
        if not HAS_MARKITDOWN:
            messagebox.showwarning(
                "Dependencia faltante",
                "markitdown no está instalado.\n→ pip install markitdown[all]",
            )

    # ── Apertura de archivo ───────────────────────────────────────────────────

    def _browse(self, _=None):
        path = dialog_open()
        if path:
            self._path_var.set(path)
            self._convert(path)

    def _entry_go(self, _=None):
        raw = self._path_var.get().strip()
        if raw:
            self._convert(self._normalize(raw))

    def _entry_paste(self, _=None):
        """Convierte automáticamente si se pega una ruta válida o URI."""
        # El contenido real del portapapeles llega tras el evento Paste,
        # así que lo procesamos con un after() corto.
        self.root.after(50, self._check_pasted)

    def _check_pasted(self):
        raw = self._path_var.get().strip()
        if not raw:
            return
        path = self._normalize(raw)
        if os.path.isfile(path):
            self._convert(path)

    @staticmethod
    def _normalize(raw: str) -> str:
        """Convierte ruta plana o file:// URI a ruta de sistema."""
        raw = raw.strip().strip("\"'")
        if raw.startswith("file://"):
            parsed = urlparse(raw)
            return url2pathname(unquote(parsed.path))
        return raw

    # ── Conversión ────────────────────────────────────────────────────────────

    def _convert(self, path: str):
        if not HAS_MARKITDOWN:
            messagebox.showerror("Error", "markitdown no está instalado.\n→ pip install markitdown[all]")
            return
        if not os.path.isfile(path):
            messagebox.showerror("Archivo no encontrado", f"{path}")
            return

        ext = os.path.splitext(path)[1].lower()
        if ext and ext not in SUPPORTED_EXT:
            if not messagebox.askyesno(
                "Formato desconocido",
                f"La extensión '{ext}' puede no estar soportada.\n¿Intentar de todas formas?",
            ):
                return

        self.source_filename = os.path.splitext(os.path.basename(path))[0]
        self._status.set(f"Convirtiendo: {os.path.basename(path)} …")
        self._progress.start(10)
        self._set_btns(tk.DISABLED)

        threading.Thread(target=self._worker, args=(path,), daemon=True).start()

    def _worker(self, path: str):
        # ⚠ Hilo secundario — NUNCA tocar Tkinter aquí.
        try:
            result = MarkItDown().convert(path)
            self._q.put(("ok", result.text_content))
        except Exception as exc:
            self._q.put(("err", str(exc)))

    def _poll(self):
        """Lee la cola desde el hilo principal cada 50 ms."""
        try:
            while True:
                tag, payload = self._q.get_nowait()
                if tag == "ok":
                    self._on_ok(payload)
                else:
                    self._on_err(payload)
        except queue.Empty:
            pass
        self.root.after(50, self._poll)

    def _on_ok(self, text: str):
        self.markdown_content = text
        self._progress.stop()
        self._set_text(text)
        chars, lines = len(text), text.count("\n")
        self._status.set(f"✅ Listo — {chars:,} caracteres · {lines:,} líneas")
        self._set_btns(tk.NORMAL)

    def _on_err(self, msg: str):
        self._progress.stop()
        self._status.set("❌ Error en la conversión.")
        messagebox.showerror("Error de conversión", f"No se pudo convertir:\n\n{msg}")

    # ── Acciones ──────────────────────────────────────────────────────────────

    def _save(self):
        if not self.markdown_content:
            return
        initial = f"{self.source_filename or 'output'}.md"
        path = dialog_save(initial)
        if not path:
            return
        try:
            with open(path, "w", encoding="utf-8") as f:
                f.write(self.markdown_content)
            self._status.set(f"✅ Guardado en: {path}")
            messagebox.showinfo("Guardado", f"Archivo guardado:\n{path}")
        except Exception as exc:
            messagebox.showerror("Error al guardar", str(exc))

    def _copy(self):
        if self.markdown_content:
            self.root.clipboard_clear()
            self.root.clipboard_append(self.markdown_content)
            self._status.set("📋 Copiado al portapapeles.")

    def _clear(self):
        self.markdown_content = ""
        self.source_filename  = ""
        self._path_var.set("")
        self._set_text("")
        self._status.set("Listo.")
        self._set_btns(tk.DISABLED)

    # ── Helpers ───────────────────────────────────────────────────────────────

    def _set_text(self, text: str):
        self._text.config(state=tk.NORMAL)
        self._text.delete("1.0", tk.END)
        self._text.insert("1.0", text)
        self._text.config(state=tk.DISABLED)

    def _set_btns(self, state):
        for b in (self._btn_save, self._btn_copy, self._btn_clear):
            b.config(state=state)


# ── Entry point ───────────────────────────────────────────────────────────────

def main():
    root = tk.Tk()   # Tkinter estándar — sin XInitThreads, sin XCB crash
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()