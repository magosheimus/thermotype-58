"""
ThermoType 58 - Editor de texto e impressão direta para impressoras térmicas 58mm
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD
from PIL import Image, ImageTk, ImageOps
import os
import sys
import json
import ctypes


def resource_path(relative_path):
    """Resolve caminho do recurso tanto em desenvolvimento quanto no executável PyInstaller."""
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)
from image_processor import ImageProcessor
from printer_handler import PrinterHandler, WIN32_AVAILABLE
from text_editor import (
    list_system_fonts, load_font_history, save_font_history,
    add_to_font_history, build_font_list, clean_font_name,
    render_text_to_image, render_rich_text_to_image,
    PAPER_WIDTH_PX as TE_PAPER_WIDTH, PADDING_PX as TE_PADDING,
)


def load_icon_image(icon_name, size=(32, 32)):
    """Carrega um ícone .ico e retorna como PhotoImage"""
    try:
        path = resource_path(icon_name)
        if os.path.exists(path):
            img = Image.open(path)
            img = img.resize(size, Image.Resampling.LANCZOS)
            return ImageTk.PhotoImage(img)
    except:
        pass
    return None


def enable_dpi_awareness():
    """Habilita DPI awareness para evitar fonte borrada no Windows"""
    try:
        # Tenta configurar DPI awareness (Windows 8.1+)
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except:
        try:
            # Fallback para Windows Vista/7
            ctypes.windll.user32.SetProcessDPIAware()
        except:
            pass


class TopStartThermalApp:
    def __init__(self, root):
        self.root = root
        self.root.title("ThermoType 58")
        self.root.geometry("1000x820")
        self.root.minsize(900, 750)  # Tamanho mínimo para acomodar 2 colunas
        self.root.configure(bg="#3165c4")  # Azul para combinar com a borda
        
        # Configurar estilos vintage Windows
        self.setup_styles()
        
        # Definir ícone da janela (title bar, taskbar e todos os Toplevels futuros)
        try:
            _ico_path = resource_path("printer.ico")
            if os.path.exists(_ico_path):
                _ico_img16 = Image.open(_ico_path).resize((16, 16), Image.Resampling.LANCZOS)
                _ico_img32 = Image.open(_ico_path).resize((32, 32), Image.Resampling.LANCZOS)
                _ico_photo16 = ImageTk.PhotoImage(_ico_img16)
                _ico_photo32 = ImageTk.PhotoImage(_ico_img32)
                # wm_iconphoto(True, ...) define o ícone padrão para a raiz e todos os Toplevels
                self.root.wm_iconphoto(True, _ico_photo32, _ico_photo16)
                self.root._icon_photo_refs = (_ico_photo16, _ico_photo32)
                # iconbitmap com after() garante que a barra de título receba o .ico correto
                self.root.after(50, lambda p=_ico_path: self.root.iconbitmap(p))
        except:
            pass
        
        # Configurações
        self.PAPER_WIDTH_MM = 58
        self.DPI = 203
        self.PIXELS_PER_MM = 8
        self.PAPER_WIDTH_PX = 384  # Reduzido para compensar margens da impressora térmica
        
        # Estado
        self.original_image = None
        self.processed_image = None
        self.current_file = None
        self._current_mode = "text"  # "image" ou "text"
        self.auto_top_fix = tk.BooleanVar(value=True)
        self.manual_offset = tk.IntVar(value=0)
        self.num_copies = tk.IntVar(value=1)
        self.selected_printer = tk.StringVar()
        self.history_file = "history.json"
        self.image_history = self.load_history()
        self.thumbnail_buttons = []

        # Pilha de undo/redo para operações de imagem
        # Cada item: {"file": path_ou_None, "image": PIL.Image, "auto_top_fix": bool, "offset": int}
        self._undo_stack: list = []
        self._redo_stack: list = []
        self._MAX_UNDO = 20

        # Estado do editor de texto
        self._editor_settings_file = "editor_settings.json"
        _es = self._load_editor_settings()
        self.font_var   = tk.StringVar(value=_es.get("font", "Arial"))
        self.size_var   = tk.IntVar(value=_es.get("size", 24))
        self.bold_var   = tk.BooleanVar(value=_es.get("bold", False))
        self.italic_var = tk.BooleanVar(value=_es.get("italic", False))
        self.uline_var  = tk.BooleanVar(value=_es.get("underline", False))
        self.align_var  = tk.StringVar(value=_es.get("align", "center"))
        # Restaurar configurações persistentes de impressão
        self.auto_top_fix.set(_es.get("auto_top_fix", True))
        self.manual_offset.set(_es.get("manual_offset", 0))
        self.num_copies.set(_es.get("num_copies", 1))
        self._text_preview_job = None
        self._all_fonts = []
        self._font_history = load_font_history()
        # Armazena tags de formatação por seleção: tag_name -> {family,size,bold,italic,underline}
        self._fmt_tags: dict = {}

        # Templates de texto
        self._templates_file = "templates.json"
        self._templates = self._load_templates()

        # Carregar ícones
        self.icon_printer_header = load_icon_image("printer.ico", (48, 48))
        self.icon_printer = load_icon_image("printer.ico", (32, 32))
        self.icon_printer_small = load_icon_image("printer.ico", (24, 24))
        self.icon_open_folder = None
        
        # Processadores
        self.image_processor = ImageProcessor(self.PAPER_WIDTH_PX, self.PIXELS_PER_MM)
        self.printer_handler = PrinterHandler()
        
        self.setup_ui()
        self.refresh_printers()
        
        # Atalho: Enter para imprimir (só fora do editor de texto)
        self.root.bind('<Return>', lambda e: self.print_image() if self.root.focus_get() is not self.text_box else None)
        # Atalhos Undo/Redo globais para imagens (apenas quando o foco não está no text_box)
        self.root.bind('<Control-z>', lambda e: self.undo_image() if self.root.focus_get() is not self.text_box else None)
        self.root.bind('<Control-y>', lambda e: self.redo_image() if self.root.focus_get() is not self.text_box else None)
        self.root.bind('<Control-Z>', lambda e: self.undo_image() if self.root.focus_get() is not self.text_box else None)
        self.root.bind('<Control-Y>', lambda e: self.redo_image() if self.root.focus_get() is not self.text_box else None)
    
    def setup_styles(self):
        """Configura estilos personalizados para a interface - Tema Vintage Windows"""
        style = ttk.Style()
        style.theme_use('clam')
        
        # Estilo para botões principais - Windows 95/98 Blue
        style.configure('Primary.TButton',
                       background='#316ac5',
                       foreground='#ffffff',
                       borderwidth=2,
                       focuscolor='#003c74',
                       padding=8,
                       relief='raised',
                       font=('MS Sans Serif', 9, 'bold'))
        style.map('Primary.TButton',
                 background=[('active', '#3e85f7'), ('pressed', '#003c74'), ('disabled', '#d3d0c7')],
                 foreground=[('disabled', '#868479')],
                 relief=[('pressed', 'sunken')])
        
        # Estilo para botão secundário - Windows Classic Gray
        style.configure('Secondary.TButton',
                       background='#ece9d8',
                       foreground='#000000',
                       borderwidth=2,
                       focuscolor='#aba798',
                       padding=8,
                       relief='raised',
                       font=('MS Sans Serif', 9))
        style.map('Secondary.TButton',
                 background=[('active', '#ffffff'), ('pressed', '#d3d0c7'), ('disabled', '#d3d0c7')],
                 foreground=[('disabled', '#868479')],
                 relief=[('pressed', 'sunken')])
        
        # Estilo para frames - Windows Classic Beige
        style.configure('Card.TFrame',
                       background='#f0eee4',
                       relief='flat')
        
        # Estilo para LabelFrame - Classic Windows com borda azul
        style.configure('TLabelframe',
                       background='#f0eee4',
                       borderwidth=3,
                       relief='ridge')
        style.configure('TLabelframe.Label',
                       background='#f0eee4',
                       foreground='#215dc6',
                       font=('MS Sans Serif', 9, 'bold'))
        
        # Estilo para Labels
        style.configure('TLabel',
                       background='#f0eee4',
                       foreground='#000000',
                       font=('MS Sans Serif', 8))
        
        # Estilo para Frames genéricos
        style.configure('TFrame',
                       background='#f0eee4')
        
        # Estilo para Spinbox
        style.configure('TSpinbox',
                       background='#ffffff',
                       foreground='#000000',
                       fieldbackground='#ffffff',
                       borderwidth=1,
                       font=('MS Sans Serif', 8))
    
    def setup_ui(self):
        """Configura a interface do usuário"""
        # ── Header ────────────────────────────────────────────────────────
        header_frame = tk.Frame(self.root, bg="#3165c4", relief='flat', borderwidth=0)
        header_frame.pack(fill=tk.X, side=tk.TOP)

        header_content = tk.Frame(header_frame, bg="#3165c4", pady=10)
        header_content.pack()

        icon_label = tk.Label(
            header_content,
            image=self.icon_printer_header if self.icon_printer_header else None,
            text="🖨️" if not self.icon_printer_header else "",
            font=("Segoe UI Emoji", 32),
            bg="#3165c4", fg="#ffffff"
        )
        icon_label.pack(side=tk.LEFT, padx=(10, 10))

        text_frame = tk.Frame(header_content, bg="#3165c4")
        text_frame.pack(side=tk.LEFT)

        tk.Label(text_frame, text="ThermoType 58",
                 font=("MS Sans Serif", 18, "bold"),
                 bg="#3165c4", fg="#ffffff").pack(anchor=tk.W)
        tk.Label(text_frame,
                 text="Editor de texto e impressão direta\npara impressoras térmicas 58mm",
                 font=("MS Sans Serif", 8),
                 bg="#3165c4", fg="#ffffff", justify=tk.LEFT).pack(anchor=tk.W)

        # ── Body ──────────────────────────────────────────────────────────
        main_outer_frame = tk.Frame(self.root, bg="#3165c4", borderwidth=3, relief='raised')
        main_outer_frame.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)

        main_frame = ttk.Frame(main_outer_frame, padding="12", style='Card.TFrame')
        main_frame.pack(fill=tk.BOTH, expand=True)

        columns_frame = ttk.Frame(main_frame, style='Card.TFrame')
        columns_frame.pack(fill=tk.BOTH, expand=True)

        # ════════════════════════════════════════════════════════════════
        # COLUNA ESQUERDA (largura fixa, pack LEFT)
        # ════════════════════════════════════════════════════════════════
        left_column = ttk.Frame(columns_frame, style='Card.TFrame', width=290)
        left_column.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 8))
        left_column.pack_propagate(False)

        # ── Botão abrir imagem ─────────────────────────────────────────
        open_btn = ttk.Button(
            left_column, text="📂  Abrir Imagem",
            command=self.open_image, style='Secondary.TButton'
        )
        open_btn.pack(fill=tk.X, pady=(0, 2))

        # ── Undo / Redo ────────────────────────────────────────────────
        undo_redo_row = ttk.Frame(left_column, style='Card.TFrame')
        undo_redo_row.pack(fill=tk.X, pady=(0, 4))
        self._btn_undo = ttk.Button(
            undo_redo_row, text="↩ Desfazer",
            command=self.undo_image, style='Secondary.TButton', state=tk.DISABLED
        )
        self._btn_undo.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 2))
        self._btn_redo = ttk.Button(
            undo_redo_row, text="↪ Refazer",
            command=self.redo_image, style='Secondary.TButton', state=tk.DISABLED
        )
        self._btn_redo.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # ── Histórico de imagens ───────────────────────────────────────
        self.history_outer = ttk.Frame(left_column, style='Card.TFrame')
        self.history_outer.pack(fill=tk.X, pady=(0, 4))
        self._build_history_thumbnails(self.history_outer)

        # ── Templates de texto ────────────────────────────────────────
        self._template_name_var = tk.StringVar()
        tmpl_lf = ttk.LabelFrame(left_column, text="Templates", padding="4")
        tmpl_lf.pack(fill=tk.X, pady=(0, 4))
        self._template_combo = ttk.Combobox(
            tmpl_lf, textvariable=self._template_name_var,
            state="normal", font=('MS Sans Serif', 8)
        )
        self._template_combo.pack(fill=tk.X, pady=(0, 3))
        self._reload_template_combo()
        tmpl_btn_row = ttk.Frame(tmpl_lf)
        tmpl_btn_row.pack(fill=tk.X)
        ttk.Button(tmpl_btn_row, text="💾 Salvar",
                   command=self._template_save,
                   style='Secondary.TButton').pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 2))
        ttk.Button(tmpl_btn_row, text="📂 Carregar",
                   command=self._template_load,
                   style='Secondary.TButton').pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 2))
        ttk.Button(tmpl_btn_row, text="🗑",
                   command=self._template_delete,
                   style='Secondary.TButton', width=2).pack(side=tk.LEFT)

        # ── Controles de imagem ────────────────────────────────────────
        self.controls_frame = ttk.LabelFrame(left_column, text="Controles", padding="4")
        self.controls_frame.pack(fill=tk.X, pady=(0, 4))

        auto_fix_check = ttk.Checkbutton(
            self.controls_frame,
            text="Auto Top Fix (remover margem superior)",
            variable=self.auto_top_fix,
            command=lambda: (self._push_undo(), self.update_preview(), self._save_editor_settings())
        )
        auto_fix_check.pack(anchor=tk.W, pady=2)

        offset_frame = ttk.Frame(self.controls_frame)
        offset_frame.pack(fill=tk.X, pady=2)
        ttk.Label(offset_frame, text="Offset manual (mm):").pack(side=tk.LEFT, padx=(0, 4))
        offset_spin = ttk.Spinbox(offset_frame, from_=-50, to=50,
                    textvariable=self.manual_offset, width=8,
                    command=lambda: (self._push_undo(), self.update_preview(), self._save_editor_settings()))
        offset_spin.pack(side=tk.LEFT)
        offset_spin.bind("<FocusOut>", lambda e: (self.update_preview(), self._save_editor_settings()))

        copies_frame = ttk.Frame(self.controls_frame)
        copies_frame.pack(fill=tk.X, pady=2)
        ttk.Label(copies_frame, text="Cópias:").pack(side=tk.LEFT, padx=(0, 4))
        copies_spin = ttk.Spinbox(copies_frame, from_=1, to=100,
                    textvariable=self.num_copies, width=8,
                    command=self._save_editor_settings)
        copies_spin.pack(side=tk.LEFT)
        copies_spin.bind("<FocusOut>", lambda e: self._save_editor_settings())

        # ── Impressora ─────────────────────────────────────────────────
        printer_lf = ttk.LabelFrame(left_column, text="Impressora", padding="4")
        printer_lf.pack(fill=tk.X, pady=(0, 4))

        printer_row = ttk.Frame(printer_lf)
        printer_row.pack(fill=tk.X)
        self.printer_combo = ttk.Combobox(
            printer_row, textvariable=self.selected_printer,
            state='readonly', font=('MS Sans Serif', 8)
        )
        self.printer_combo.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))
        ttk.Button(printer_row, text="🔄", command=self.refresh_printers, width=3).pack(side=tk.LEFT)

        # ── Informações ────────────────────────────────────────────────
        info_frame = ttk.LabelFrame(left_column, text="Informações", padding="4")
        info_frame.pack(fill=tk.X, pady=(0, 4))
        self.info_label = tk.Label(
            info_frame, text="Editor de texto ativo",
            font=("MS Sans Serif", 8), bg="#f0eee4", fg="#000000",
            justify=tk.LEFT, anchor=tk.W, wraplength=240
        )
        self.info_label.pack(fill=tk.X)

        # ── Botões de ação (parte inferior da coluna) ──────────────────
        self.print_btn = ttk.Button(
            left_column, text="🖨️ IMPRIMIR",
            command=self.print_image, style='Primary.TButton'
        )
        self.print_btn.pack(side=tk.BOTTOM, fill=tk.X, pady=(4, 0))

        self.save_text_btn = ttk.Button(
            left_column, text="💾 Salvar Imagem",
            command=self._text_save_image, style='Secondary.TButton'
        )
        self.save_text_btn.pack(side=tk.BOTTOM, fill=tk.X, pady=(0, 2))

        self.edit_text_btn = ttk.Button(
            left_column, text="✏️ Editar Texto",
            command=self._show_text_mode, style='Secondary.TButton'
        )
        # Aparece apenas quando uma imagem está carregada no preview

        # ════════════════════════════════════════════════════════════════
        # COLUNA DIREITA (ocupa o restante, pack LEFT + expand)
        # ════════════════════════════════════════════════════════════════
        right_column = ttk.Frame(columns_frame, style='Card.TFrame')
        right_column.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        right_column.rowconfigure(1, weight=1)
        right_column.columnconfigure(0, weight=1)

        # ── Toolbar de formatação (sempre visível) ───────────────────
        self.toolbar_frame = ttk.LabelFrame(right_column, text="Formatação", padding=4)
        self.toolbar_frame.grid(row=0, column=0, sticky='ew', pady=(0, 4))
        self._build_text_toolbar(self.toolbar_frame)

        # ── Área principal: editor de texto / preview de imagem ──────
        self.preview_label_frame = ttk.LabelFrame(
            right_column, text="Editor de Texto", padding="5"
        )
        self.preview_label_frame.grid(row=1, column=0, sticky='nsew')

        preview_border = tk.Frame(
            self.preview_label_frame, bg="#b0b0b0", relief='sunken', borderwidth=2
        )
        preview_border.pack(fill=tk.BOTH, expand=True, padx=3, pady=3)

        # Canvas (preview de imagem) — oculto por padrão
        self.preview_canvas = tk.Canvas(
            preview_border, bg="#ffffff",
            width=550, height=500,
            highlightthickness=0, relief='flat'
        )
        self.preview_scroll = ttk.Scrollbar(
            preview_border, orient="vertical",
            command=self.preview_canvas.yview
        )
        self.preview_canvas.configure(yscrollcommand=self.preview_scroll.set)
        self.preview_canvas.bind('<Configure>', self._on_canvas_configure)
        self.preview_canvas.drop_target_register(DND_FILES)
        self.preview_canvas.dnd_bind('<<Drop>>', self.on_drop)
        # NÃO empacotado agora — aparece ao carregar imagem

        # ── Tampo cinza + papel branco na largura exata da impressora ────
        # A largura do papel em tela = (364 px impressora / 203 DPI) × DPI tela
        # Como fontes usam pontos (unit. independente de DPI), as quebras de
        # linha que o usuário vê são idênticas às da impressão.
        self.text_editor_desk = tk.Frame(preview_border, bg="#f0eee4")
        self.text_editor_desk.pack(fill=tk.BOTH, expand=True)
        self.text_editor_desk.bind('<Configure>', self._on_desk_configure)

        # "Papel" — largura fixa proporcional ao papel físico (sem scrollbar visível)
        self.text_area_frame = tk.Frame(
            self.text_editor_desk, bg="#ffffff",
            relief='solid', borderwidth=1
        )
        self.text_area_frame.pack_propagate(False)

        self.text_box = tk.Text(
            self.text_area_frame,
            font=("Arial", 14),
            bg="#ffffff", fg="#000000",
            relief="flat", borderwidth=0,
            wrap="word", undo=True,
        )
        self.text_box.pack(fill=tk.BOTH, expand=True)
        self.text_box.bind("<KeyRelease>", lambda e: self._on_text_key_release())
        self.text_box.drop_target_register(DND_FILES)
        self.text_box.dnd_bind('<<Drop>>', self.on_drop)

        # Placeholder (comportamento nativo)
        self._placeholder_text = "Digite seu texto aqui..."
        self._placeholder_active = True
        self.text_box.tag_configure("placeholder", foreground="#aaaaaa")
        self.text_box.insert("1.0", self._placeholder_text, "placeholder")
        self.text_box.bind("<FocusIn>",  self._on_textbox_focus_in)
        self.text_box.bind("<FocusOut>", self._on_textbox_focus_out)

        # Info de dimensões do texto renderizado
        self.text_preview_info = tk.Label(
            self.preview_label_frame, text="", font=("MS Sans Serif", 8),
            bg="#f0eee4", fg="#555555"
        )
        self.text_preview_info.pack(pady=(1, 0))

        # Inicializar lista de fontes (lazy, em background)
        self.root.after(200, self._init_fonts)
        self.root.after(150, self._on_desk_configure)

    # ── Construção auxiliar da toolbar ────────────────────────────────────

    def _build_text_toolbar(self, parent):
        """Constrói toolbar de formatação de texto dentro de `parent` — tudo em uma linha."""
        BG = "#f0eee4"

        row = ttk.Frame(parent)
        row.pack(fill=tk.X)

        ttk.Label(row, text="Fonte:").pack(side=tk.LEFT, padx=(0, 3))
        self.font_combo = ttk.Combobox(
            row, textvariable=self.font_var,
            width=18, font=("MS Sans Serif", 8), state="normal"
        )
        self.font_combo.pack(side=tk.LEFT, padx=(0, 6))
        self.font_combo.bind("<<ComboboxSelected>>", self._on_font_selected)
        self.font_combo.bind("<Return>", self._on_font_selected)
        # Ao abrir o dropdown, rolar sempre para o topo (onde está o histórico)
        self.font_combo.bind("<ButtonPress-1>",
            lambda e: self.root.after(20, self._scroll_font_combo_top))

        ttk.Label(row, text="Tam:").pack(side=tk.LEFT, padx=(0, 3))
        sp = ttk.Spinbox(
            row, from_=6, to=200, textvariable=self.size_var,
            width=4, command=self._schedule_text_preview
        )
        sp.pack(side=tk.LEFT, padx=(0, 8))
        sp.bind("<Return>", lambda e: self._schedule_text_preview())
        sp.bind("<FocusOut>", lambda e: self._schedule_text_preview())
        sp.bind("<KeyRelease>", lambda e: self._schedule_text_preview())

        # Separador visual
        ttk.Separator(row, orient="vertical").pack(side=tk.LEFT, fill=tk.Y, padx=(0, 6))

        style_f = ttk.Frame(row)
        style_f.pack(side=tk.LEFT, padx=(0, 8))
        tk.Checkbutton(style_f, text="N", font=("MS Sans Serif", 9, "bold"),
                       variable=self.bold_var, bg=BG,
                       command=self._schedule_text_preview,
                       indicatoron=False, relief="raised", width=2, padx=2
                       ).pack(side=tk.LEFT, padx=1)
        tk.Checkbutton(style_f, text="I", font=("MS Sans Serif", 9, "italic"),
                       variable=self.italic_var, bg=BG,
                       command=self._schedule_text_preview,
                       indicatoron=False, relief="raised", width=2, padx=2
                       ).pack(side=tk.LEFT, padx=1)
        tk.Checkbutton(style_f, text="S̲", font=("MS Sans Serif", 9),
                       variable=self.uline_var, bg=BG,
                       command=self._schedule_text_preview,
                       indicatoron=False, relief="raised", width=2, padx=2
                       ).pack(side=tk.LEFT, padx=1)

        # Separador visual
        ttk.Separator(row, orient="vertical").pack(side=tk.LEFT, fill=tk.Y, padx=(0, 6))

        ttk.Label(row, text="Alinhamento:").pack(side=tk.LEFT, padx=(0, 3))
        for lbl, val in [("Esq", "left"), ("Cen", "center"), ("Dir", "right")]:
            tk.Radiobutton(
                row, text=lbl, variable=self.align_var, value=val,
                bg=BG, font=("MS Sans Serif", 8),
                command=self._schedule_text_preview,
                indicatoron=False, relief="raised", padx=4, pady=2
            ).pack(side=tk.LEFT, padx=1)

    def _build_history_thumbnails(self, parent):
        """Popula o frame de miniaturas de histórico."""
        for w in parent.winfo_children():
            w.destroy()
        if not self.image_history:
            return
        hf = ttk.LabelFrame(parent, text="Imagens Recentes", padding="5")
        hf.pack(fill=tk.X)
        tc = ttk.Frame(hf)
        tc.pack(fill=tk.X)
        self.thumbnail_buttons = []
        for img_path in self.image_history[:5]:
            if os.path.exists(img_path):
                try:
                    img = Image.open(img_path)
                    img.thumbnail((60, 60), Image.Resampling.LANCZOS)
                    photo = ImageTk.PhotoImage(img)
                    btn = tk.Button(
                        tc, image=photo,
                        command=lambda p=img_path: self.load_image_from_path(p),
                        relief='raised', borderwidth=2, bg='#ffffff', cursor='hand2'
                    )
                    btn.image = photo
                    btn.pack(side=tk.LEFT, padx=2, pady=2)
                    self.thumbnail_buttons.append(btn)
                except Exception:
                    pass

    def _scroll_font_combo_top(self):
        """Força o dropdown do combobox de fontes a rolar para o topo (histórico)."""
        try:
            # Acessa o Listbox interno do ttk.Combobox via Tk
            popdown = self.font_combo.tk.eval(
                f'ttk::combobox::PopdownWindow {self.font_combo}'
            )
            self.font_combo.tk.eval(f'{popdown}.f.l yview moveto 0')
        except Exception:
            pass

    # ── Placeholder do editor ──────────────────────────────────────────────

    def _on_textbox_focus_in(self, event=None):
        if self._placeholder_active:
            self.text_box.delete("1.0", "end")
            self.text_box.tag_remove("placeholder", "1.0", "end")
            self._placeholder_active = False
            self._apply_textbox_font()

    def _on_textbox_focus_out(self, event=None):
        if not self.text_box.get("1.0", "end-1c").strip():
            self._placeholder_active = True
            self.text_box.insert("1.0", self._placeholder_text, "placeholder")

    def _get_real_text(self):
        """Retorna o texto real (ignora placeholder)."""
        if self._placeholder_active:
            return ""
        return self.text_box.get("1.0", "end-1c").strip()

    # ── Templates de texto ────────────────────────────────────────────────

    def _load_templates(self) -> list:
        try:
            if os.path.exists(self._templates_file):
                with open(self._templates_file, "r", encoding="utf-8") as f:
                    return json.load(f)
        except Exception:
            pass
        return []

    def _save_templates_to_file(self):
        try:
            with open(self._templates_file, "w", encoding="utf-8") as f:
                json.dump(self._templates, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def _reload_template_combo(self):
        if hasattr(self, '_template_combo'):
            self._template_combo["values"] = [t["name"] for t in self._templates]

    def _template_save(self):
        """Salva o texto e formatação atual como template nomeado."""
        text = self._get_real_text()
        if not text:
            messagebox.showwarning("Aviso", "Digite algum texto antes de salvar o template!")
            return
        name = self._template_name_var.get().strip()
        if not name:
            messagebox.showwarning("Aviso", "Digite um nome para o template no campo acima!")
            return
        template = {
            "name":      name,
            "text":      text,
            "font":      self.font_var.get(),
            "size":      self.size_var.get(),
            "bold":      self.bold_var.get(),
            "italic":    self.italic_var.get(),
            "underline": self.uline_var.get(),
            "align":     self.align_var.get(),
        }
        # Substituir se já existe com o mesmo nome, senão inserir no topo
        self._templates = [t for t in self._templates if t["name"] != name]
        self._templates.insert(0, template)
        self._save_templates_to_file()
        self._reload_template_combo()
        messagebox.showinfo("Template salvo", f"Template '{name}' salvo com sucesso!")

    def _template_load(self):
        """Carrega o template selecionado, restaurando texto e formatação."""
        name = self._template_name_var.get().strip()
        if not name:
            messagebox.showwarning("Aviso", "Selecione ou digite o nome de um template!")
            return
        tmpl = next((t for t in self._templates if t["name"] == name), None)
        if tmpl is None:
            messagebox.showwarning("Aviso", f"Template '{name}' não encontrado.")
            return
        # Aplicar formatação
        self.font_var.set(tmpl.get("font", "Arial"))
        self.size_var.set(tmpl.get("size", 24))
        self.bold_var.set(tmpl.get("bold", False))
        self.italic_var.set(tmpl.get("italic", False))
        self.uline_var.set(tmpl.get("underline", False))
        self.align_var.set(tmpl.get("align", "center"))
        # Carregar texto no editor
        self.text_box.delete("1.0", "end")
        self.text_box.tag_remove("placeholder", "1.0", "end")
        self._placeholder_active = False
        self.text_box.insert("1.0", tmpl["text"])
        self._fmt_tags.clear()
        self._apply_textbox_font()
        self._schedule_text_preview()
        # Garantir modo texto visível
        if not self.text_editor_desk.winfo_ismapped():
            self._show_text_mode()

    def _template_delete(self):
        """Exclui o template selecionado após confirmação."""
        name = self._template_name_var.get().strip()
        if not name or not any(t["name"] == name for t in self._templates):
            return
        if not messagebox.askyesno("Excluir Template", f"Excluir o template '{name}'?"):
            return
        self._templates = [t for t in self._templates if t["name"] != name]
        self._save_templates_to_file()
        self._template_name_var.set("")
        self._reload_template_combo()

    # ── Configurações do editor (persistência) ────────────────────────────

    def _load_editor_settings(self) -> dict:
        try:
            if os.path.exists(self._editor_settings_file):
                with open(self._editor_settings_file, "r", encoding="utf-8") as f:
                    return json.load(f)
        except Exception:
            pass
        return {}

    def _save_editor_settings(self):
        try:
            data = {
                "font":         self.font_var.get(),
                "size":         self.size_var.get(),
                "bold":         self.bold_var.get(),
                "italic":       self.italic_var.get(),
                "underline":    self.uline_var.get(),
                "align":        self.align_var.get(),
                "printer":      self.selected_printer.get(),
                "auto_top_fix": self.auto_top_fix.get(),
                "manual_offset": self.manual_offset.get(),
                "num_copies":   self.num_copies.get(),
            }
            with open(self._editor_settings_file, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    # ── Fontes (lazy init) ─────────────────────────────────────────────────

    def _init_fonts(self):
        """Carrega lista de fontes do sistema (chamado após UI já renderizada)."""
        self._all_fonts = list_system_fonts()
        if not self.font_var.get() or self.font_var.get() not in self._all_fonts:
            default = "Arial" if "Arial" in self._all_fonts else (self._all_fonts[0] if self._all_fonts else "Arial")
            self.font_var.set(default)
        self._reload_font_combo()
        self._apply_textbox_font()

    def _reload_font_combo(self):
        self.font_combo["values"] = build_font_list(self._all_fonts, self._font_history)

    def _on_font_selected(self, event=None):
        raw = self.font_var.get()
        if raw.startswith("─"):
            return
        clean = clean_font_name(raw)
        self.font_var.set(clean)
        self._font_history = add_to_font_history(clean, self._font_history)
        save_font_history(self._font_history)
        self._reload_font_combo()
        self._apply_textbox_font()
        self._schedule_text_preview()

    def _apply_textbox_font(self):
        """Aplica fonte à seleção (se houver) ou a todo o widget."""
        try:
            family = self.font_var.get() or "Arial"
            size = max(6, self.size_var.get())
            bold = self.bold_var.get()
            italic = self.italic_var.get()
            uline = self.uline_var.get()
            styles = (["bold"] if bold else []) + (["italic"] if italic else [])
            font_spec = (family, size) + tuple(styles)

            # Verificar se há texto selecionado
            try:
                sel_start = self.text_box.index("sel.first")
                sel_end   = self.text_box.index("sel.last")
                has_sel = True
            except tk.TclError:
                has_sel = False

            if has_sel:
                # Aplicar tag apenas na seleção
                safe_family = family.replace(" ", "_").replace("-", "_")
                tag_name = f"fmt_{safe_family}_{size}_{'B' if bold else 'n'}{'I' if italic else 'n'}{'U' if uline else 'n'}"
                if tag_name not in self._fmt_tags:
                    self._fmt_tags[tag_name] = {
                        "family": family, "size": size,
                        "bold": bold, "italic": italic, "underline": uline,
                    }
                    self.text_box.tag_configure(tag_name, font=font_spec, underline=uline)
                # Remover outros fmt tags da região selecionada (evita conflito)
                for other in list(self._fmt_tags):
                    if other != tag_name:
                        self.text_box.tag_remove(other, sel_start, sel_end)
                self.text_box.tag_add(tag_name, sel_start, sel_end)
            else:
                # Sem seleção → mudança global: limpar fmt tags e atualizar fonte base
                for tag in list(self._fmt_tags):
                    self.text_box.tag_remove(tag, "1.0", "end")
                self._fmt_tags.clear()
                self.text_box.config(font=font_spec)

            # Alinhamento (sempre global)
            align = self.align_var.get()
            for a in ("left", "center", "right"):
                self.text_box.tag_configure(f"align_{a}", justify=a)
                self.text_box.tag_remove(f"align_{a}", "1.0", "end")
            self.text_box.tag_add(f"align_{align}", "1.0", "end")
        except Exception:
            pass

    # ── Preview de texto ───────────────────────────────────────────────────

    def _on_text_key_release(self):
        """Reaplica alinhamento a cada tecla e agenda o preview."""
        align = self.align_var.get()
        self.text_box.tag_add(f"align_{align}", "1.0", "end")
        self._schedule_text_preview()

    def _schedule_text_preview(self, _event=None):
        self._apply_textbox_font()
        self._save_editor_settings()
        if self._text_preview_job:
            self.root.after_cancel(self._text_preview_job)
        self._text_preview_job = self.root.after(300, self._update_text_preview)

    def _get_formatted_lines(self):
        """Extrai linhas com formatação do widget Text.
        Retorna [(line_text, fmt_dict), ...] onde fmt_dict tem:
        family, size, bold, italic, underline, align.
        """
        text = self._get_real_text()
        if not text:
            return []
        lines = text.split('\n')
        base_align = self.align_var.get()
        base_fmt = {
            "family":    clean_font_name(self.font_var.get()).strip() or "Arial",
            "size":      max(6, self.size_var.get()),
            "bold":      self.bold_var.get(),
            "italic":    self.italic_var.get(),
            "underline": self.uline_var.get(),
            "align":     base_align,
        }
        if not self._fmt_tags:
            return [(line, base_fmt) for line in lines]

        # Offset de char para o início de cada linha
        line_starts = []
        pos = 0
        for line in lines:
            line_starts.append(pos)
            pos += len(line) + 1  # +1 pelo \n

        # Converter índice Tk "linha.col" (1-based) para offset de char
        def tk_to_char(tk_idx):
            r, c = map(int, str(tk_idx).split('.'))
            r -= 1  # 0-based
            if r < 0 or r >= len(line_starts):
                return 0
            return line_starts[r] + c

        # Coletar ranges de cada fmt tag
        tagged = []  # (char_start, char_end, fmt)
        for tag_name, fmt in self._fmt_tags.items():
            ranges = self.text_box.tag_ranges(tag_name)
            for i in range(0, len(ranges), 2):
                cs = tk_to_char(ranges[i])
                ce = tk_to_char(ranges[i + 1])
                full_fmt = dict(fmt)
                full_fmt["align"] = base_align
                tagged.append((cs, ce, full_fmt))

        # Para cada linha, usar o fmt que cobre mais caracteres
        result = []
        for idx, line in enumerate(lines):
            l_start = line_starts[idx]
            l_end   = l_start + len(line)
            best_fmt = base_fmt
            best_cov = 0
            for cs, ce, tfmt in tagged:
                cov = max(0, min(l_end, ce) - max(l_start, cs))
                if cov > best_cov:
                    best_cov = cov
                    best_fmt = tfmt
            result.append((line, best_fmt))
        return result

    def _update_text_preview(self):
        lines = self._get_formatted_lines()
        if not lines:
            self.text_preview_info.config(text="")
            return
        try:
            img = render_rich_text_to_image(
                formatted_lines=lines,
                paper_width=TE_PAPER_WIDTH,
                padding=TE_PADDING,
            )
            self._rendered_text_image = img
            h_mm = img.height / 8
            self.text_preview_info.config(
                text=f"{TE_PAPER_WIDTH}×{img.height}px  |  58×{h_mm:.1f}mm"
            )
        except Exception as e:
            print(f"[preview] {e}")

    def _get_rendered_text_image(self):
        """Renderiza o texto com formatação por seleção e retorna a imagem PIL."""
        lines = self._get_formatted_lines()
        if not lines:
            return None
        return render_rich_text_to_image(
            formatted_lines=lines,
            paper_width=TE_PAPER_WIDTH,
            padding=TE_PADDING,
        )

    def _text_save_image(self):
        """Salva o texto renderizado como imagem."""
        img = self._get_rendered_text_image()
        if img is None:
            messagebox.showwarning("Aviso", "Digite algum texto antes de salvar!")
            return
        path = filedialog.asksaveasfilename(
            title="Salvar imagem de texto",
            defaultextension=".png",
            filetypes=[("PNG", "*.png"), ("JPEG", "*.jpg"), ("BMP", "*.bmp"), ("Todos", "*.*")],
            initialfile="texto_impresso.png",
        )
        if not path:
            return
        try:
            ext = os.path.splitext(path)[1].lower()
            if ext in (".jpg", ".jpeg"):
                img.convert("RGB").save(path, "JPEG", quality=95)
            else:
                img.save(path)
            messagebox.showinfo("Salvo", f"Imagem salva em:\n{path}")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao salvar:\n{e}")

    # ── Canvas helpers ─────────────────────────────────────────────────────

    def _on_canvas_configure(self, event=None):
        if self.original_image:
            self.update_preview()
        else:
            self._draw_dotted_border()
        
    def _draw_image_placeholder(self):
        self.preview_canvas.delete("all")
        cw = self.preview_canvas.winfo_width() or 550
        ch = self.preview_canvas.winfo_height() or 500
        self.preview_canvas.create_text(
            cw // 2, ch // 2,
            text="Arraste uma imagem aqui ou clique em 'Abrir Imagem'",
            fill="#888888", font=("MS Sans Serif", 9),
            tags="placeholder"
        )
    
    def _draw_dotted_border(self, event=None):
        """Desenha borda pontilhada no canvas de preview"""
        if not hasattr(self, 'preview_canvas'):
            return
        width = self.preview_canvas.winfo_width()
        height = self.preview_canvas.winfo_height()
        self.preview_canvas.delete("dotted_border")
        margin = 10
        self.preview_canvas.create_rectangle(
            margin, margin, width - margin, height - margin,
            outline="#888888", width=2, dash=(5, 5), tags="dotted_border"
        )
    
    def refresh_printers(self):
        """Atualiza lista de impressoras disponíveis no sistema"""
        printers = self.printer_handler.list_printers()
        
        if not printers:
            printers = ["Nenhuma impressora encontrada"]
        
        self.printer_combo['values'] = printers
        
        current = self.selected_printer.get()
        if current in printers:
            return  # Manter seleção atual

        # 1. Tentar restaurar a última impressora usada (salva nas configurações)
        last_printer = self._load_editor_settings().get("printer", "")
        if last_printer and last_printer in printers:
            self.selected_printer.set(last_printer)
            return

        # 2. Tentar selecionar a impressora padrão do sistema
        default_selected = False
        if WIN32_AVAILABLE:
            try:
                import win32print
                default = win32print.GetDefaultPrinter()
                if default in printers:
                    self.selected_printer.set(default)
                    default_selected = True
            except Exception:
                pass

        if not default_selected:
            self.selected_printer.set(printers[0])
    
    def load_history(self):
        """Carrega histórico de imagens do arquivo JSON"""
        try:
            if os.path.exists(self.history_file):
                with open(self.history_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
        except:
            pass
        return []
    
    def save_history(self):
        """Salva histórico de imagens no arquivo JSON"""
        try:
            with open(self.history_file, 'w', encoding='utf-8') as f:
                json.dump(self.image_history, f, ensure_ascii=False, indent=2)
        except:
            pass
    
    def add_to_history(self, file_path):
        """Adiciona arquivo ao histórico (máximo 10 itens)"""
        # Remover se já existe
        if file_path in self.image_history:
            self.image_history.remove(file_path)
        
        # Adicionar no início
        self.image_history.insert(0, file_path)
        
        # Limitar a 10 itens
        self.image_history = self.image_history[:10]
        
        # Salvar
        self.save_history()
        
        # Atualizar combobox se existir
        if hasattr(self, 'history_combo'):
            self.history_combo['values'] = self.image_history
    
    def load_image_from_path(self, path):
        """Carrega imagem a partir de um caminho"""
        if path and os.path.exists(path):
            self.load_image(path)
        else:
            messagebox.showerror("Erro", "Arquivo não encontrado!")
    
    def on_drop(self, event):
        """Handler para drag and drop"""
        # Obter o caminho do arquivo
        file_path = event.data
        
        # Remover chaves {} se existirem (Windows)
        if file_path.startswith('{') and file_path.endswith('}'):
            file_path = file_path[1:-1]
        
        # Remover espaços extras
        file_path = file_path.strip()
        
        # Carregar a imagem
        if os.path.isfile(file_path):
            self.load_image(file_path)
    
    def open_image(self):
        """Abre diálogo para selecionar imagem"""
        file_path = filedialog.askopenfilename(
            title="Selecione uma imagem",
            filetypes=[
                ("Imagens", "*.png *.jpg *.jpeg *.bmp"),
                ("PNG", "*.png"),
                ("JPEG", "*.jpg *.jpeg"),
                ("BMP", "*.bmp"),
                ("Todos os arquivos", "*.*")
            ]
        )
        
        if file_path:
            self.load_image(file_path)
    
    def load_image(self, file_path):
        """Carrega e processa a imagem, mostrando o canvas no lugar da textarea."""
        try:
            # Salvar texto atual no histórico antes de sair do modo texto
            if not self.original_image:
                self._save_text_to_history()
            self._push_undo()
            self.current_file = file_path
            self.original_image = Image.open(file_path)
            self.add_to_history(file_path)
            self.process_image()
            self._current_mode = "image"
            # Substituir textarea pelo canvas
            self.text_editor_desk.pack_forget()
            self.text_preview_info.pack_forget()
            self.preview_scroll.pack(side=tk.RIGHT, fill=tk.Y)
            self.preview_canvas.pack(padx=3, pady=3, fill=tk.BOTH, expand=True)
            self.preview_label_frame.config(text="Preview da Impressão")
            # Mostrar botão para voltar ao editor de texto
            self.edit_text_btn.pack(side=tk.BOTTOM, fill=tk.X, pady=(0, 2),
                                    before=self.save_text_btn)
            self.update_preview()
            self._build_history_thumbnails(self.history_outer)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar imagem: {str(e)}")
    
    def _save_text_to_history(self):
        """Renderiza o texto atual e salva no histórico de miniaturas."""
        img = self._get_rendered_text_image()
        if img is None:
            return
        try:
            cache_dir = os.path.dirname(os.path.abspath(__file__))
            thumb_path = os.path.join(cache_dir, "_last_text_preview.png")
            img.save(thumb_path)
            # Inserir no topo do histórico (remove duplicata se já existir)
            if thumb_path in self.image_history:
                self.image_history.remove(thumb_path)
            self.image_history.insert(0, thumb_path)
            self.image_history = self.image_history[:10]
            self.save_history()
            self._build_history_thumbnails(self.history_outer)
        except Exception:
            pass

    def _show_text_mode(self):
        """Volta para o editor de texto, ocultando o preview de imagem."""
        self._current_mode = "text"
        self.preview_canvas.pack_forget()
        self.preview_scroll.pack_forget()
        self.text_editor_desk.pack(fill=tk.BOTH, expand=True)
        self.text_preview_info.pack(pady=(1, 0))
        self.preview_label_frame.config(text="Editor de Texto")
        self.edit_text_btn.pack_forget()
        self.root.after(10, self._on_desk_configure)
        self.text_box.focus_set()

    def _on_desk_configure(self, event=None):
        """Centraliza e dimensiona o 'papel' para corresponder à largura física
        do papel 58 mm — quebras de linha no editor == quebras na impressora."""
        if not hasattr(self, 'text_editor_desk'):
            return
        if not self.text_editor_desk.winfo_ismapped():
            return
        try:
            # Largura útil da impressora em polegadas → pixels de tela
            screen_dpi = self.root.winfo_fpixels('1i')
            if screen_dpi < 72:   # sanidade
                screen_dpi = 96
            printer_usable_px = TE_PAPER_WIDTH - 2 * TE_PADDING  # 364
            paper_screen_w = round(screen_dpi * printer_usable_px / 203)
            scrollbar_w = 20          # largura da scrollbar
            border_w = 2              # borderwidth=1 × 2 lados
            frame_w = paper_screen_w + scrollbar_w + border_w

            desk_w = self.text_editor_desk.winfo_width()
            desk_h = self.text_editor_desk.winfo_height()
            if desk_w < 10 or desk_h < 10:
                return

            pad = 10
            x = max(pad, (desk_w - frame_w) // 2)
            actual_w = min(frame_w, desk_w - 2 * pad)
            actual_h = max(desk_h - 2 * pad, 50)

            self.text_area_frame.place(
                x=x, y=pad, width=actual_w, height=actual_h
            )
        except Exception:
            self.text_area_frame.place(
                relx=0.5, y=10, anchor='n',
                relwidth=0.98, relheight=0.96
            )

    def process_image(self):
        """Processa a imagem aplicando correções"""
        if not self.original_image:
            return
        
        # Redimensionar para largura correta (464px)
        self.processed_image = self.image_processor.resize_to_width(self.original_image)
        
        # Aplicar auto top fix se habilitado
        if self.auto_top_fix.get():
            self.processed_image = self.image_processor.remove_top_margin(self.processed_image)
        
        # Aplicar offset manual
        offset_mm = self.manual_offset.get()
        if offset_mm != 0:
            self.processed_image = self.image_processor.apply_offset(
                self.processed_image,
                offset_mm
            )

    # ── Undo / Redo ────────────────────────────────────────────────────────

    def _push_undo(self):
        """Empurra o estado atual na pilha de undo antes de fazer uma mudança."""
        if self.original_image is None:
            return
        state = {
            "file":         self.current_file,
            "image":        self.original_image.copy(),
            "auto_top_fix": self.auto_top_fix.get(),
            "offset":       self.manual_offset.get(),
        }
        self._undo_stack.append(state)
        if len(self._undo_stack) > self._MAX_UNDO:
            self._undo_stack.pop(0)
        self._redo_stack.clear()
        self._update_undo_buttons()

    def _restore_state(self, state):
        """Restaura um estado salvo na pilha."""
        self.current_file     = state["file"]
        self.original_image   = state["image"].copy()
        self.auto_top_fix.set(state["auto_top_fix"])
        self.manual_offset.set(state["offset"])
        # Mostrar canvas de imagem
        self.text_editor_desk.pack_forget()
        self.text_preview_info.pack_forget()
        self.preview_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.preview_canvas.pack(padx=3, pady=3, fill=tk.BOTH, expand=True)
        self.preview_label_frame.config(text="Preview da Impressão")
        self._current_mode = "image"
        self.edit_text_btn.pack(side=tk.BOTTOM, fill=tk.X, pady=(0, 2),
                                before=self.save_text_btn)
        self.update_preview()
        self._update_undo_buttons()

    def undo_image(self):
        """Desfaz a última operação de imagem."""
        if not self._undo_stack:
            return
        # Salvar estado atual em redo
        if self.original_image is not None:
            redo_state = {
                "file":         self.current_file,
                "image":        self.original_image.copy(),
                "auto_top_fix": self.auto_top_fix.get(),
                "offset":       self.manual_offset.get(),
            }
            self._redo_stack.append(redo_state)
        self._restore_state(self._undo_stack.pop())

    def redo_image(self):
        """Refaz a última operação desfeita."""
        if not self._redo_stack:
            return
        if self.original_image is not None:
            undo_state = {
                "file":         self.current_file,
                "image":        self.original_image.copy(),
                "auto_top_fix": self.auto_top_fix.get(),
                "offset":       self.manual_offset.get(),
            }
            self._undo_stack.append(undo_state)
        self._restore_state(self._redo_stack.pop())

    def _update_undo_buttons(self):
        """Atualiza estado (enabled/disabled) dos botões de undo/redo."""
        if hasattr(self, '_btn_undo'):
            self._btn_undo.config(state=tk.NORMAL if self._undo_stack else tk.DISABLED)
        if hasattr(self, '_btn_redo'):
            self._btn_redo.config(state=tk.NORMAL if self._redo_stack else tk.DISABLED)
    
    def update_preview(self):
        """Atualiza o preview da imagem no canvas."""
        if not self.original_image:
            return
        self.process_image()
        self.preview_canvas.delete("all")

        canvas_width = self.preview_canvas.winfo_width()
        canvas_height = self.preview_canvas.winfo_height()
        if canvas_width <= 1:
            canvas_width = 550
        if canvas_height <= 1:
            canvas_height = 500

        preview_image = self.processed_image.copy()
        preview_image.thumbnail((canvas_width - 40, canvas_height - 40), Image.Resampling.LANCZOS)
        self.photo_preview = ImageTk.PhotoImage(preview_image)

        center_x = canvas_width // 2
        start_y = 30

        self.preview_canvas.create_image(center_x, start_y, anchor=tk.N, image=self.photo_preview)

        margin = 20
        self.preview_canvas.create_line(
            margin, start_y, canvas_width - margin, start_y,
            fill="red", width=2, dash=(5, 5)
        )
        self.preview_canvas.create_text(
            center_x, start_y - 20,
            text="INÍCIO DA IMPRESSÃO (Y=0)",
            fill="red", font=("MS Sans Serif", 7, "bold"), anchor=tk.N
        )

        height_mm = self.processed_image.height / self.PIXELS_PER_MM
        self.info_label.config(
            text=f"Dimensões: {self.processed_image.width}x{self.processed_image.height}px\n"
                 f"Altura: {height_mm:.1f}mm | Largura: {self.PAPER_WIDTH_MM}mm"
        )
        self.preview_canvas.configure(scrollregion=(
            0, 0, canvas_width, start_y + preview_image.height + 20
        ))
    
    def print_image(self):
        """Envia imagem carregada ou texto digitado para impressão."""
        printer_name = self.selected_printer.get()
        if printer_name and printer_name != "Nenhuma impressora encontrada":
            self.printer_handler.set_printer(printer_name)
            self._save_editor_settings()  # persistir última impressora usada

        num_copies = self.num_copies.get()

        if self._current_mode == "image" and self.original_image and self.processed_image:
            # ── Imprimir imagem carregada ─────────────────────────────────
            try:
                print_img = self.image_processor.convert_to_monochrome(self.processed_image)
                ok = sum(1 for _ in range(num_copies)
                         if self.printer_handler.print_image(print_img))
                if ok == num_copies:
                    messagebox.showinfo("Sucesso", f"{num_copies} cópia(s) enviada(s) com sucesso!")
                elif ok > 0:
                    messagebox.showwarning("Parcial", f"{ok} de {num_copies} cópia(s) impressa(s)")
                else:
                    messagebox.showerror("Erro", "Falha ao enviar para impressora")
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao imprimir: {str(e)}")
        else:
            # ── Imprimir texto digitado ───────────────────────────────────
            img = self._get_rendered_text_image()
            if img is None:
                messagebox.showwarning("Aviso", "Digite algum texto ou abra uma imagem antes de imprimir!")
                return
            try:
                ok = sum(1 for _ in range(num_copies)
                         if self.printer_handler.print_image(img.convert("1")))
                if ok == num_copies:
                    messagebox.showinfo("Sucesso", f"{num_copies} cópia(s) enviada(s) com sucesso!")
                    self._save_text_to_history()
                elif ok > 0:
                    messagebox.showwarning("Parcial", f"{ok} de {num_copies} cópia(s) impressa(s)")
                    self._save_text_to_history()
                else:
                    messagebox.showerror("Erro", "Falha ao enviar para impressora")
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao imprimir:\n{e}")
    


def main():
    # Habilitar DPI awareness ANTES de criar a janela Tkinter
    enable_dpi_awareness()
    
    root = TkinterDnD.Tk()
    app = TopStartThermalApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
