#!/usr/bin/env python3
"""
Handwriting Recognition to Excel - Desktop Interface (Tkinter)
===============================================================
Native desktop GUI for uploading handwritten images and exporting
the recognized content as formatted Excel spreadsheets.

Run with:
    python app_desktop.py
"""

import os
import sys
import threading
import tempfile
import logging
from pathlib import Path
from tkinter import (
    Tk, Frame, Label, Button, StringVar, BooleanVar, OptionMenu,
    filedialog, messagebox, ttk, Canvas, Scrollbar, BOTH, LEFT,
    RIGHT, TOP, BOTTOM, X, Y, W, E, N, S, HORIZONTAL, VERTICAL,
    WORD, END, DISABLED, NORMAL,
)
from tkinter.font import Font as TkFont

from PIL import Image, ImageTk

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Color scheme
# ---------------------------------------------------------------------------

COLORS = {
    "bg": "#f1f5f9",
    "card": "#ffffff",
    "primary": "#1e3a5f",
    "primary_hover": "#2d5a8a",
    "accent": "#2d6a9f",
    "success": "#059669",
    "success_bg": "#ecfdf5",
    "warning": "#d97706",
    "error": "#dc2626",
    "text": "#1e293b",
    "text_secondary": "#64748b",
    "border": "#e2e8f0",
    "header_bg": "#1e3a5f",
    "header_text": "#ffffff",
    "btn_process": "#2563eb",
    "btn_process_hover": "#1d4ed8",
    "btn_download": "#059669",
    "btn_download_hover": "#047857",
}


# ---------------------------------------------------------------------------
# Application
# ---------------------------------------------------------------------------

class HandwritingApp:
    def __init__(self):
        self.root = Tk()
        self.root.title("Reconhecimento de Escrita Manual para Excel")
        self.root.geometry("960x720")
        self.root.minsize(800, 600)
        self.root.configure(bg=COLORS["bg"])

        # State
        self.image_path = None
        self.output_path = None
        self.excel_data = None
        self.table_data = None
        self.processing = False

        # Variables
        self.var_engine = StringVar(value="both")
        self.var_enhance = StringVar(value="standard")
        self.var_mode = StringVar(value="auto")
        self.var_header = BooleanVar(value=True)
        self.var_status = StringVar(value="Aguardando imagem...")

        # Fonts
        self.font_title = TkFont(family="Segoe UI", size=18, weight="bold")
        self.font_subtitle = TkFont(family="Segoe UI", size=11)
        self.font_heading = TkFont(family="Segoe UI", size=13, weight="bold")
        self.font_body = TkFont(family="Segoe UI", size=10)
        self.font_small = TkFont(family="Segoe UI", size=9)
        self.font_btn = TkFont(family="Segoe UI", size=11, weight="bold")

        self._build_ui()

    def _build_ui(self):
        """Build the complete UI."""
        # Header
        self._build_header()

        # Main content
        main_frame = Frame(self.root, bg=COLORS["bg"])
        main_frame.pack(fill=BOTH, expand=True, padx=20, pady=(10, 20))

        # Left panel - Upload & Preview
        left = Frame(main_frame, bg=COLORS["bg"])
        left.pack(side=LEFT, fill=BOTH, expand=True, padx=(0, 10))

        self._build_upload_section(left)
        self._build_preview_section(left)

        # Right panel - Settings & Results
        right = Frame(main_frame, bg=COLORS["bg"], width=320)
        right.pack(side=RIGHT, fill=Y, padx=(10, 0))
        right.pack_propagate(False)

        self._build_settings_section(right)
        self._build_actions_section(right)
        self._build_results_section(right)

        # Status bar
        self._build_status_bar()

    def _build_header(self):
        """Build application header."""
        header = Frame(self.root, bg=COLORS["header_bg"], height=80)
        header.pack(fill=X)
        header.pack_propagate(False)

        inner = Frame(header, bg=COLORS["header_bg"])
        inner.pack(fill=BOTH, expand=True, padx=25, pady=15)

        Label(
            inner,
            text="Reconhecimento de Escrita Manual",
            font=self.font_title,
            bg=COLORS["header_bg"],
            fg=COLORS["header_text"],
        ).pack(anchor=W)

        Label(
            inner,
            text="Envie uma foto de escrita manual e receba uma planilha Excel organizada",
            font=self.font_subtitle,
            bg=COLORS["header_bg"],
            fg="#b8d4e8",
        ).pack(anchor=W)

    def _build_upload_section(self, parent):
        """Build the image upload section."""
        card = Frame(parent, bg=COLORS["card"], highlightbackground=COLORS["border"],
                     highlightthickness=1, padx=20, pady=20)
        card.pack(fill=X, pady=(0, 10))

        Label(
            card, text="Enviar Imagem", font=self.font_heading,
            bg=COLORS["card"], fg=COLORS["text"],
        ).pack(anchor=W, pady=(0, 10))

        # Drop zone
        self.drop_zone = Frame(
            card, bg="#f8fafc", highlightbackground="#cbd5e1",
            highlightthickness=2, height=80, cursor="hand2",
        )
        self.drop_zone.pack(fill=X)
        self.drop_zone.pack_propagate(False)

        self.lbl_drop = Label(
            self.drop_zone,
            text="Clique aqui para selecionar uma imagem\n(JPG, PNG, BMP, TIFF)",
            font=self.font_body, bg="#f8fafc", fg=COLORS["text_secondary"],
            cursor="hand2",
        )
        self.lbl_drop.pack(expand=True)

        # Bind click events
        for widget in (self.drop_zone, self.lbl_drop):
            widget.bind("<Button-1>", lambda e: self._select_file())

    def _build_preview_section(self, parent):
        """Build the image preview section."""
        card = Frame(parent, bg=COLORS["card"], highlightbackground=COLORS["border"],
                     highlightthickness=1, padx=20, pady=20)
        card.pack(fill=BOTH, expand=True)

        Label(
            card, text="Pre-visualizacao", font=self.font_heading,
            bg=COLORS["card"], fg=COLORS["text"],
        ).pack(anchor=W, pady=(0, 10))

        # Image canvas
        canvas_frame = Frame(card, bg=COLORS["card"])
        canvas_frame.pack(fill=BOTH, expand=True)

        self.canvas = Canvas(canvas_frame, bg="#f8fafc", highlightthickness=0)
        self.canvas.pack(fill=BOTH, expand=True)

        self.lbl_no_image = Label(
            self.canvas, text="Nenhuma imagem carregada",
            font=self.font_body, bg="#f8fafc", fg=COLORS["text_secondary"],
        )
        self.lbl_no_image.place(relx=0.5, rely=0.5, anchor="center")

        self._photo_ref = None  # Keep reference to prevent GC

    def _build_settings_section(self, parent):
        """Build the settings panel."""
        card = Frame(parent, bg=COLORS["card"], highlightbackground=COLORS["border"],
                     highlightthickness=1, padx=15, pady=15)
        card.pack(fill=X, pady=(0, 10))

        Label(
            card, text="Configuracoes", font=self.font_heading,
            bg=COLORS["card"], fg=COLORS["text"],
        ).pack(anchor=W, pady=(0, 10))

        # Engine
        self._add_setting(card, "Motor OCR:", self.var_engine, {
            "both": "Ambos (melhor)",
            "tesseract": "Tesseract",
            "easyocr": "EasyOCR",
        })

        # Enhance
        self._add_setting(card, "Melhoria:", self.var_enhance, {
            "standard": "Padrao",
            "aggressive": "Agressivo",
        })

        # Mode
        self._add_setting(card, "Deteccao:", self.var_mode, {
            "auto": "Automatico",
            "table": "Tabela",
            "lines": "Linhas",
        })

        # Header checkbox
        chk_frame = Frame(card, bg=COLORS["card"])
        chk_frame.pack(fill=X, pady=3)
        ttk.Checkbutton(
            chk_frame, text="1a linha e cabecalho",
            variable=self.var_header,
        ).pack(anchor=W)

    def _add_setting(self, parent, label_text, var, options):
        """Add a labeled dropdown setting."""
        frame = Frame(parent, bg=COLORS["card"])
        frame.pack(fill=X, pady=3)

        Label(
            frame, text=label_text, font=self.font_small,
            bg=COLORS["card"], fg=COLORS["text_secondary"],
        ).pack(anchor=W)

        # Create OptionMenu with display values
        display_to_value = {v: k for k, v in options.items()}
        value_to_display = options

        display_var = StringVar(value=value_to_display[var.get()])

        def on_change(*args):
            var.set(display_to_value[display_var.get()])

        display_var.trace_add("write", on_change)

        menu = OptionMenu(frame, display_var, *value_to_display.values())
        menu.configure(
            font=self.font_small, bg="white", fg=COLORS["text"],
            activebackground=COLORS["accent"], activeforeground="white",
            highlightthickness=0, relief="solid", borderwidth=1,
        )
        menu.pack(fill=X)

    def _build_actions_section(self, parent):
        """Build the action buttons."""
        card = Frame(parent, bg=COLORS["card"], highlightbackground=COLORS["border"],
                     highlightthickness=1, padx=15, pady=15)
        card.pack(fill=X, pady=(0, 10))

        # Process button
        self.btn_process = Button(
            card,
            text="Processar Imagem",
            font=self.font_btn,
            bg=COLORS["btn_process"],
            fg="white",
            activebackground=COLORS["btn_process_hover"],
            activeforeground="white",
            relief="flat",
            cursor="hand2",
            height=2,
            command=self._process_image,
            state=DISABLED,
        )
        self.btn_process.pack(fill=X, pady=(0, 8))

        # Progress bar
        self.progress = ttk.Progressbar(card, mode="determinate", length=280)
        self.progress.pack(fill=X, pady=(0, 5))

        self.lbl_progress = Label(
            card, text="", font=self.font_small,
            bg=COLORS["card"], fg=COLORS["text_secondary"],
        )
        self.lbl_progress.pack(anchor=W)

        # Download button
        self.btn_download = Button(
            card,
            text="Baixar Excel",
            font=self.font_btn,
            bg=COLORS["btn_download"],
            fg="white",
            activebackground=COLORS["btn_download_hover"],
            activeforeground="white",
            relief="flat",
            cursor="hand2",
            height=2,
            command=self._save_excel,
            state=DISABLED,
        )
        self.btn_download.pack(fill=X, pady=(5, 0))

    def _build_results_section(self, parent):
        """Build the results display."""
        card = Frame(parent, bg=COLORS["card"], highlightbackground=COLORS["border"],
                     highlightthickness=1, padx=15, pady=15)
        card.pack(fill=BOTH, expand=True)

        Label(
            card, text="Resultado", font=self.font_heading,
            bg=COLORS["card"], fg=COLORS["text"],
        ).pack(anchor=W, pady=(0, 10))

        # Stats
        self.lbl_stats = Label(
            card, text="Nenhum resultado ainda",
            font=self.font_body, bg=COLORS["card"], fg=COLORS["text_secondary"],
            wraplength=270, justify=LEFT,
        )
        self.lbl_stats.pack(anchor=W, pady=(0, 10))

        # Preview text area with scrollbar
        preview_frame = Frame(card, bg=COLORS["card"])
        preview_frame.pack(fill=BOTH, expand=True)

        scrollbar = Scrollbar(preview_frame, orient=VERTICAL)
        scrollbar.pack(side=RIGHT, fill=Y)

        from tkinter import Text
        self.txt_preview = Text(
            preview_frame, font=self.font_small,
            bg="#f8fafc", fg=COLORS["text"],
            wrap=WORD, height=10,
            yscrollcommand=scrollbar.set,
            state=DISABLED, relief="solid", borderwidth=1,
        )
        self.txt_preview.pack(fill=BOTH, expand=True)
        scrollbar.config(command=self.txt_preview.yview)

    def _build_status_bar(self):
        """Build the bottom status bar."""
        bar = Frame(self.root, bg=COLORS["border"], height=30)
        bar.pack(fill=X, side=BOTTOM)
        bar.pack_propagate(False)

        Label(
            bar, textvariable=self.var_status,
            font=self.font_small, bg=COLORS["border"], fg=COLORS["text_secondary"],
        ).pack(side=LEFT, padx=15)

    # -----------------------------------------------------------------------
    # Actions
    # -----------------------------------------------------------------------

    def _select_file(self):
        """Open file dialog to select an image."""
        filetypes = [
            ("Imagens", "*.jpg *.jpeg *.png *.bmp *.tiff *.tif *.webp"),
            ("JPEG", "*.jpg *.jpeg"),
            ("PNG", "*.png"),
            ("Todos os arquivos", "*.*"),
        ]

        path = filedialog.askopenfilename(
            title="Selecionar imagem com escrita manual",
            filetypes=filetypes,
        )

        if path:
            self.image_path = path
            self._load_preview(path)
            self.btn_process.config(state=NORMAL)
            self.var_status.set(f"Imagem carregada: {Path(path).name}")
            self.lbl_drop.config(
                text=f"Selecionado: {Path(path).name}\n(clique para trocar)",
                fg=COLORS["accent"],
            )

    def _load_preview(self, path):
        """Load and display image preview in the canvas."""
        try:
            self.lbl_no_image.place_forget()

            img = Image.open(path)

            # Update canvas and get dimensions
            self.canvas.update_idletasks()
            cw = max(self.canvas.winfo_width(), 200)
            ch = max(self.canvas.winfo_height(), 200)

            # Scale image to fit canvas
            ratio = min(cw / img.width, ch / img.height)
            new_w = int(img.width * ratio)
            new_h = int(img.height * ratio)
            img_resized = img.resize((new_w, new_h), Image.LANCZOS)

            self._photo_ref = ImageTk.PhotoImage(img_resized)

            self.canvas.delete("all")
            self.canvas.create_image(cw // 2, ch // 2, image=self._photo_ref, anchor="center")

        except Exception as e:
            logger.error(f"Preview error: {e}")
            self.lbl_no_image.place(relx=0.5, rely=0.5, anchor="center")
            self.lbl_no_image.config(text=f"Erro ao carregar: {e}")

    def _process_image(self):
        """Start image processing in a background thread."""
        if self.processing or not self.image_path:
            return

        self.processing = True
        self.btn_process.config(state=DISABLED, text="Processando...")
        self.btn_download.config(state=DISABLED)
        self.progress["value"] = 0

        thread = threading.Thread(target=self._process_worker, daemon=True)
        thread.start()

    def _process_worker(self):
        """Background worker for image processing."""
        try:
            self._update_progress(10, "Carregando imagem...")

            from handwriting_to_excel import HandwritingToExcel, ImagePreprocessor

            self._update_progress(20, "Pre-processando imagem...")

            # Create temp output
            tmp_output = tempfile.mktemp(suffix=".xlsx")

            self._update_progress(40, "Executando OCR...")

            pipeline = HandwritingToExcel(
                engine=self.var_engine.get(),
                languages=["pt", "en"],
                enhance_level=self.var_enhance.get(),
                mode=self.var_mode.get(),
            )

            self._update_progress(60, "Detectando estrutura...")

            table_data = pipeline.convert(
                image_path=self.image_path,
                output_path=tmp_output,
                first_row_header=self.var_header.get(),
            )

            self._update_progress(85, "Gerando Excel...")

            # Read Excel bytes
            with open(tmp_output, "rb") as f:
                self.excel_data = f.read()

            self.table_data = table_data
            self.output_path = tmp_output

            try:
                os.unlink(tmp_output)
            except OSError:
                pass

            self._update_progress(100, "Concluido!")

            # Update UI on main thread
            self.root.after(0, self._on_process_complete)

        except Exception as e:
            logger.exception("Processing error")
            self.root.after(0, lambda: self._on_process_error(str(e)))

    def _update_progress(self, value, text):
        """Thread-safe progress update."""
        self.root.after(0, lambda: self._set_progress(value, text))

    def _set_progress(self, value, text):
        self.progress["value"] = value
        self.lbl_progress.config(text=text)
        self.var_status.set(text)

    def _on_process_complete(self):
        """Handle successful processing."""
        self.processing = False
        self.btn_process.config(state=NORMAL, text="Processar Imagem")
        self.btn_download.config(state=NORMAL)

        if self.table_data:
            num_rows = len(self.table_data)
            num_cols = max(len(r) for r in self.table_data) if self.table_data else 0

            self.lbl_stats.config(
                text=f"Reconhecimento concluido!\n\n"
                     f"Linhas detectadas: {num_rows}\n"
                     f"Colunas detectadas: {num_cols}",
                fg=COLORS["success"],
            )

            # Fill preview
            self.txt_preview.config(state=NORMAL)
            self.txt_preview.delete("1.0", END)

            for i, row in enumerate(self.table_data[:50]):
                line = " | ".join(str(cell) for cell in row)
                self.txt_preview.insert(END, f"Ln {i+1}: {line}\n")

            if len(self.table_data) > 50:
                self.txt_preview.insert(END, f"\n... +{len(self.table_data) - 50} linhas")

            self.txt_preview.config(state=DISABLED)

            self.var_status.set(
                f"Pronto! {num_rows} linhas x {num_cols} colunas detectadas"
            )
        else:
            self.lbl_stats.config(
                text="Nenhum texto detectado.\nTente modo 'Agressivo'.",
                fg=COLORS["warning"],
            )

    def _on_process_error(self, error_msg):
        """Handle processing error."""
        self.processing = False
        self.btn_process.config(state=NORMAL, text="Processar Imagem")
        self.progress["value"] = 0
        self.lbl_progress.config(text="")
        self.var_status.set(f"Erro: {error_msg}")

        messagebox.showerror(
            "Erro no Processamento",
            f"Ocorreu um erro ao processar a imagem:\n\n{error_msg}\n\n"
            "Dicas:\n"
            "- Verifique se o Tesseract esta instalado\n"
            "- Tente usar apenas EasyOCR no motor\n"
            "- Verifique a qualidade da imagem",
        )

    def _save_excel(self):
        """Save the generated Excel file."""
        if not self.excel_data:
            return

        default_name = Path(self.image_path).stem + ".xlsx" if self.image_path else "resultado.xlsx"

        path = filedialog.asksaveasfilename(
            title="Salvar planilha Excel",
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[
                ("Excel", "*.xlsx"),
                ("Todos os arquivos", "*.*"),
            ],
        )

        if path:
            try:
                with open(path, "wb") as f:
                    f.write(self.excel_data)
                self.var_status.set(f"Arquivo salvo: {path}")
                messagebox.showinfo(
                    "Arquivo Salvo",
                    f"Planilha Excel salva com sucesso em:\n{path}",
                )
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao salvar: {e}")

    def run(self):
        """Start the application."""
        self.root.mainloop()


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    app = HandwritingApp()
    app.run()


if __name__ == "__main__":
    main()
