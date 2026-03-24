"""
ThermoType 58 - Editor de Texto
Permite criar textos formatados e imprimir diretamente na impressora térmica.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, font as tkfont
from PIL import Image, ImageDraw, ImageFont, ImageTk
import os
import sys
import json
import platform


def resource_path(relative_path):
    """Resolve caminho do recurso tanto em desenvolvimento quanto no executável PyInstaller."""
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)

# ── Constantes ──────────────────────────────────────────────────────────────
PAPER_WIDTH_PX = 384          # Largura do papel 58 mm na resolução da impressora
FONT_HISTORY_FILE = "font_history.json"
MAX_FONT_HISTORY = 8          # Quantas fontes recentes mostrar no topo da lista
PADDING_PX = 10               # Margem lateral interna do papel
PREVIEW_SCALE = 1.5           # Fator de escala do preview (para melhor legibilidade)

BG_MAIN   = "#f0eee4"
BG_HEADER = "#3165c4"
BG_WHITE  = "#ffffff"
FG_BLUE   = "#215dc6"
FONT_UI   = ("MS Sans Serif", 8)
FONT_BOLD = ("MS Sans Serif", 9, "bold")


# ── Utilidades de fonte ──────────────────────────────────────────────────────

def list_system_fonts():
    """Retorna lista de nomes de famílias de fontes instaladas no sistema."""
    families = sorted(set(tkfont.families()))
    # Filtrar fontes que começam com @ (variantes verticais do Windows)
    return [f for f in families if not f.startswith("@")]


def load_font_history():
    """Carrega histórico de fontes usadas."""
    try:
        if os.path.exists(FONT_HISTORY_FILE):
            with open(FONT_HISTORY_FILE, "r", encoding="utf-8") as fh:
                data = json.load(fh)
                return data if isinstance(data, list) else []
    except Exception:
        pass
    return []


def save_font_history(history: list):
    """Salva histórico de fontes."""
    try:
        with open(FONT_HISTORY_FILE, "w", encoding="utf-8") as fh:
            json.dump(history[:MAX_FONT_HISTORY], fh, ensure_ascii=False, indent=2)
    except Exception:
        pass


def add_to_font_history(font_name: str, history: list) -> list:
    """Adiciona fonte ao início do histórico, sem duplicatas."""
    if font_name in history:
        history.remove(font_name)
    history.insert(0, font_name)
    return history[:MAX_FONT_HISTORY]


def build_font_list(all_fonts: list, history: list) -> list:
    """
    Monta a lista de fontes para o Combobox:
    - Fontes recentes (com prefixo ★) no topo
    - Separador visual  ──────────
    - Todas as fontes em ordem alfabética
    """
    result = []
    if history:
        for f in history:
            result.append(f"★ {f}")
        result.append("─" * 28)
    result.extend(all_fonts)
    return result


def clean_font_name(display_name: str) -> str:
    """Remove prefixo '★ ' do nome exibido no combobox."""
    name = display_name.strip()
    if name.startswith("★ "):
        name = name[2:]
    return name


# ── Renderização de imagem de texto ─────────────────────────────────────────

def find_truetype_font(family: str, bold: bool = False, italic: bool = False):
    """
    Localiza o arquivo de fonte usando o Registro do Windows.
    Verifica HKLM (fontes do sistema) E HKCU (fontes instaladas pelo usuário).
    """
    import winreg

    sys_fonts_dir = os.path.join(os.environ.get("WINDIR", "C:\\Windows"), "Fonts")
    user_fonts_dir = os.path.join(
        os.environ.get("LOCALAPPDATA", ""), "Microsoft", "Windows", "Fonts"
    )
    family_lower = family.lower()

    # Preferência de estilo (da mais específica para genérica)
    if bold and italic:
        style_prefs = ["bold italic", "bolditalic", "bold"]
    elif bold:
        style_prefs = ["bold"]
    elif italic:
        style_prefs = ["italic", "oblique"]
    else:
        style_prefs = ["regular", ""]

    def _pick_best(candidates):
        """Escolhe o melhor candidato da lista [(name_lower, path)]."""
        for pref in style_prefs:
            for name, path in candidates:
                if pref and pref in name:
                    return path
        for name, path in candidates:
            if "regular" in name or not any(
                s in name for s in ["bold", "italic", "light", "thin", "black", "medium"]
            ):
                return path
        return candidates[0][1]

    def _search_reg_key(hive, default_dir):
        """Varre uma chave de registro e retorna candidatos para a família."""
        try:
            key = winreg.OpenKey(
                hive,
                r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"
            )
        except OSError:
            return []
        found = []
        i = 0
        while True:
            try:
                name, path, _ = winreg.EnumValue(key, i)
                if family_lower in name.lower():
                    if not os.path.isabs(path):
                        path = os.path.join(default_dir, path)
                    if os.path.exists(path) and path.lower().endswith((".ttf", ".otf")):
                        found.append((name.lower(), path))
                i += 1
            except OSError:
                break
        winreg.CloseKey(key)
        return found

    # 1. Tentar HKCU (fontes do usuário — instaladas sem admin)
    candidates = _search_reg_key(winreg.HKEY_CURRENT_USER, user_fonts_dir)

    # 2. Tentar HKLM (fontes do sistema)
    if not candidates:
        candidates = _search_reg_key(winreg.HKEY_LOCAL_MACHINE, sys_fonts_dir)

    if candidates:
        return _pick_best(candidates)

    # 3. Fallback: varrer as pastas de fontes diretamente
    lower_family = family_lower.replace(" ", "")
    for fonts_dir in [sys_fonts_dir, user_fonts_dir]:
        if not os.path.isdir(fonts_dir):
            continue
        for fname in os.listdir(fonts_dir):
            base, ext = os.path.splitext(fname)
            if ext.lower() not in (".ttf", ".otf"):
                continue
            base_lower = base.lower().replace(" ", "").replace("-", "").replace("_", "")
            if base_lower.startswith(lower_family) or lower_family in base_lower:
                return os.path.join(fonts_dir, fname)
    return None


def render_text_to_image(
    text: str,
    font_family: str,
    font_size: int,
    bold: bool,
    italic: bool,
    underline: bool,
    align: str,          # "left" | "center" | "right"
    paper_width: int = PAPER_WIDTH_PX,
    padding: int = PADDING_PX,
) -> Image.Image:
    """
    Renderiza o texto num Image RGB com largura = paper_width.
    font_size é em pontos tipográficos (igual ao Word/Photoshop).
    Converte para pixels a 203 DPI: px = pt × 203/72.
    """
    # Converter pontos tipográficos → pixels na resolução da impressora
    PRINTER_DPI = 203
    font_size_px = max(8, round(font_size * PRINTER_DPI / 72))
    usable_width = paper_width - 2 * padding

    # ── Tentar carregar fonte TrueType ────────────────────────────────────
    pil_font = None
    font_path = find_truetype_font(font_family, bold, italic)
    if font_path:
        try:
            pil_font = ImageFont.truetype(font_path, font_size_px)
        except Exception:
            pil_font = None

    if pil_font is None:
        # Fallback: fonte padrão do PIL (não TrueType)
        try:
            pil_font = ImageFont.load_default(size=font_size_px)
        except Exception:
            pil_font = ImageFont.load_default()

    # ── Quebra de linhas manual ───────────────────────────────────────────
    def wrap_line(line: str) -> list:
        """Quebra uma linha simples para caber em usable_width."""
        words = line.split(" ")
        wrapped = []
        current = ""
        for word in words:
            test = (current + " " + word).strip()
            bbox = pil_font.getbbox(test)
            w = bbox[2] - bbox[0]
            if w <= usable_width:
                current = test
            else:
                if current:
                    wrapped.append(current)
                # Palavra maior que a linha? quebra forçada por caractere
                while True:
                    bbox_word = pil_font.getbbox(word)
                    ww = bbox_word[2] - bbox_word[0]
                    if ww <= usable_width:
                        current = word
                        break
                    # Encontrar quantos chars cabem
                    for ch in range(len(word), 0, -1):
                        bbox_ch = pil_font.getbbox(word[:ch])
                        if bbox_ch[2] - bbox_ch[0] <= usable_width:
                            wrapped.append(word[:ch])
                            word = word[ch:]
                            break
        if current:
            wrapped.append(current)
        return wrapped if wrapped else [""]

    raw_lines = text.split("\n")
    lines = []
    for raw in raw_lines:
        lines.extend(wrap_line(raw) if raw.strip() else [""])

    # Espaçamento baseado em métricas reais (evita corte no fundo)
    try:
        ascent, descent = pil_font.getmetrics()
        line_spacing = ascent + descent + 4
    except Exception:
        line_spacing = font_size_px + 8

    # Canvas generoso: nunca cortará o texto
    total_height = line_spacing * (len(lines) + 1) + 2 * padding

    # ── Criar imagem ─────────────────────────────────────────────────────
    img = Image.new("RGB", (paper_width, max(total_height, 40)), color=(255, 255, 255))
    draw = ImageDraw.Draw(img)

    y = padding
    for line in lines:
        if not line:
            y += line_spacing
            continue

        bbox = pil_font.getbbox(line)
        text_w = bbox[2] - bbox[0]
        text_h = bbox[3] - bbox[1]

        if align == "center":
            x = (paper_width - text_w) // 2
        elif align == "right":
            x = paper_width - padding - text_w
        else:  # left
            x = padding

        draw.text((x, y), line, font=pil_font, fill=(0, 0, 0))

        # Sublinhado manual
        if underline:
            uy = y + text_h + 2
            draw.line([(x, uy), (x + text_w, uy)], fill=(0, 0, 0), width=1)

        y += line_spacing

    # ── Autocrop: remover espaço em branco inferior extra ────────────────
    img_px = img.load()
    last_row = padding
    for row in range(img.height - 1, padding - 1, -1):
        for col in range(img.width):
            if img_px[col, row] != (255, 255, 255):
                last_row = row
                break
        if last_row > padding:
            break
    final_h = min(last_row + padding + 4, img.height)
    return img.crop((0, 0, paper_width, final_h))


def render_rich_text_to_image(
    formatted_lines: list,
    paper_width: int = PAPER_WIDTH_PX,
    padding: int = PADDING_PX,
) -> Image.Image:
    """
    Renderiza texto com formatação por linha.
    formatted_lines: [(line_text, fmt_dict), ...]
    fmt_dict: {family, size, bold, italic, underline, align}
    Cada linha pode ter fonte/tamanho independente.
    """
    PRINTER_DPI = 203
    usable_width = paper_width - 2 * padding

    # Cache de fontes PIL para evitar recarregar o mesmo arquivo
    _font_cache: dict = {}

    def _get_pil_font(family, size, bold, italic):
        key = (family, size, bold, italic)
        if key not in _font_cache:
            size_px = max(8, round(size * PRINTER_DPI / 72))
            path = find_truetype_font(family, bold, italic)
            try:
                f = ImageFont.truetype(path, size_px) if path else None
            except Exception:
                f = None
            if f is None:
                try:
                    f = ImageFont.load_default(size=size_px)
                except Exception:
                    f = ImageFont.load_default()
            _font_cache[key] = f
        return _font_cache[key]

    def _wrap(text, font):
        """Quebra texto para caber em usable_width com a fonte dada."""
        if not text:
            return [""]
        words = text.split(" ")
        rows, current = [], ""
        for word in words:
            test = (current + " " + word).strip()
            w = font.getbbox(test)[2] - font.getbbox(test)[0]
            if w <= usable_width:
                current = test
            else:
                if current:
                    rows.append(current)
                while True:
                    bw = font.getbbox(word)[2] - font.getbbox(word)[0]
                    if bw <= usable_width:
                        current = word
                        break
                    for ch in range(len(word), 0, -1):
                        if font.getbbox(word[:ch])[2] - font.getbbox(word[:ch])[0] <= usable_width:
                            rows.append(word[:ch])
                            word = word[ch:]
                            break
        if current:
            rows.append(current)
        return rows if rows else [""]

    # Construir todas as linhas a renderizar: (text, font, underline, align, ascent, descent)
    render_rows = []
    for line_text, fmt in formatted_lines:
        family  = fmt.get("family", "Arial")
        size    = fmt.get("size", 14)
        bold    = fmt.get("bold", False)
        italic  = fmt.get("italic", False)
        uline   = fmt.get("underline", False)
        align   = fmt.get("align", "left")
        font    = _get_pil_font(family, size, bold, italic)
        try:
            asc, desc = font.getmetrics()
        except Exception:
            asc, desc = round(size * PRINTER_DPI / 72), max(4, round(size * PRINTER_DPI / 72) // 5)

        if not line_text:
            render_rows.append(("", font, uline, align, asc, desc))
        else:
            for sub in _wrap(line_text, font):
                render_rows.append((sub, font, uline, align, asc, desc))

    if not render_rows:
        render_rows = [("", _get_pil_font("Arial", 14, False, False), False, "left", 20, 4)]

    # Altura total
    total_h = padding + sum(asc + desc + 4 for _, _, _, _, asc, desc in render_rows) + padding

    img  = Image.new("RGB", (paper_width, max(total_h, 40)), (255, 255, 255))
    draw = ImageDraw.Draw(img)

    y = padding
    for text, font, uline, align, asc, desc in render_rows:
        line_h = asc + desc + 4
        if text:
            bbox   = font.getbbox(text)
            text_w = bbox[2] - bbox[0]
            if align == "center":
                x = (paper_width - text_w) // 2
            elif align == "right":
                x = paper_width - padding - text_w
            else:
                x = padding
            draw.text((x, y), text, font=font, fill=(0, 0, 0))
            if uline:
                draw.line([(x, y + asc + 2), (x + text_w, y + asc + 2)],
                          fill=(0, 0, 0), width=1)
        y += line_h

    # Autocrop
    img_px  = img.load()
    last_row = padding
    for row in range(img.height - 1, padding - 1, -1):
        for col in range(img.width):
            if img_px[col, row] != (255, 255, 255):
                last_row = row
                break
        if last_row > padding:
            break
    final_h = min(last_row + padding + 4, img.height)
    return img.crop((0, 0, paper_width, final_h))


# ── Janela do Editor ─────────────────────────────────────────────────────────

class TextEditorWindow:
    def __init__(self, parent, printer_handler, image_processor, selected_printer_var):
        self.parent = parent
        self.printer_handler = printer_handler
        self.image_processor = image_processor
        self.selected_printer_var = selected_printer_var

        # Dados internos
        self.font_history = load_font_history()
        self.all_fonts = list_system_fonts()
        self.rendered_image: Image.Image | None = None

        # Variáveis tkinter
        self.font_var    = tk.StringVar()
        self.size_var    = tk.IntVar(value=24)
        self.bold_var    = tk.BooleanVar(value=False)
        self.italic_var  = tk.BooleanVar(value=False)
        self.uline_var   = tk.BooleanVar(value=False)
        self.align_var   = tk.StringVar(value="left")

        # Inicializar fonte padrão
        default_font = "Arial" if "Arial" in self.all_fonts else (self.all_fonts[0] if self.all_fonts else "TkDefaultFont")
        self.font_var.set(default_font)

        self._build_window()
        self._schedule_preview_update()

    # ── Construção da janela ────────────────────────────────────────────────

    def _build_window(self):
        self.win = tk.Toplevel(self.parent)
        self.win.title("✏️ Editor de Texto — ThermoType 58")
        self.win.geometry("900x700")
        self.win.minsize(820, 580)
        self.win.configure(bg=BG_HEADER)

        try:
            _ico_path = resource_path("printer.ico")
            if os.path.exists(_ico_path):
                self.win.iconbitmap(_ico_path)
        except Exception:
            pass

        # ── Header ────────────────────────────────────────────────────────
        hdr = tk.Frame(self.win, bg=BG_HEADER)
        hdr.pack(fill=tk.X)
        tk.Label(
            hdr, text="✏️  Editor de Texto",
            font=("MS Sans Serif", 14, "bold"),
            bg=BG_HEADER, fg="#ffffff", pady=8
        ).pack(side=tk.LEFT, padx=12)

        # ── Corpo principal ────────────────────────────────────────────────
        body_outer = tk.Frame(self.win, bg=BG_HEADER, borderwidth=3, relief="raised")
        body_outer.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)

        body = ttk.Frame(body_outer, padding=10, style="Card.TFrame")
        body.pack(fill=tk.BOTH, expand=True)

        # Grid 2 colunas
        body.columnconfigure(0, weight=1)
        body.columnconfigure(1, weight=1)
        body.rowconfigure(0, weight=1)

        # ── COLUNA ESQUERDA: controles + textarea ─────────────────────────
        left = ttk.Frame(body, style="Card.TFrame")
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 8))

        self._build_toolbar(left)
        self._build_textarea(left)
        self._build_action_buttons(left)

        # ── COLUNA DIREITA: preview ────────────────────────────────────────
        right = ttk.Frame(body, style="Card.TFrame")
        right.grid(row=0, column=1, sticky="nsew")

        self._build_preview(right)

    # ── Toolbar de formatação ───────────────────────────────────────────────

    def _build_toolbar(self, parent):
        toolbar = ttk.LabelFrame(parent, text="Formatação", padding=8)
        toolbar.pack(fill=tk.X, pady=(0, 6))

        # ── Linha 1: Fonte + Tamanho ────────────────────────────────────
        row1 = ttk.Frame(toolbar)
        row1.pack(fill=tk.X, pady=(0, 4))

        ttk.Label(row1, text="Fonte:").pack(side=tk.LEFT, padx=(0, 4))

        self.font_combo = ttk.Combobox(
            row1,
            textvariable=self.font_var,
            width=22,
            font=FONT_UI,
            state="normal",
        )
        self._reload_font_combo()
        self.font_combo.pack(side=tk.LEFT, padx=(0, 8))
        self.font_combo.bind("<<ComboboxSelected>>", self._on_font_selected)
        self.font_combo.bind("<Return>", self._on_font_selected)

        ttk.Label(row1, text="Tamanho:").pack(side=tk.LEFT, padx=(0, 4))
        size_spin = ttk.Spinbox(
            row1,
            from_=6, to=200,
            textvariable=self.size_var,
            width=6,
            command=self._schedule_preview_update,
        )
        size_spin.pack(side=tk.LEFT)
        size_spin.bind("<Return>", lambda e: self._schedule_preview_update())
        size_spin.bind("<FocusOut>", lambda e: self._schedule_preview_update())

        # ── Linha 2: Estilos + Alinhamento ─────────────────────────────
        row2 = ttk.Frame(toolbar)
        row2.pack(fill=tk.X)

        # Negrito / Itálico / Sublinhado
        style_frame = ttk.Frame(row2)
        style_frame.pack(side=tk.LEFT, padx=(0, 16))

        bold_btn = tk.Checkbutton(
            style_frame, text="N", font=("MS Sans Serif", 9, "bold"),
            variable=self.bold_var, bg=BG_MAIN,
            command=self._schedule_preview_update, indicatoron=False,
            relief="raised", width=3, padx=4
        )
        bold_btn.pack(side=tk.LEFT, padx=1)

        italic_btn = tk.Checkbutton(
            style_frame, text="I", font=("MS Sans Serif", 9, "italic"),
            variable=self.italic_var, bg=BG_MAIN,
            command=self._schedule_preview_update, indicatoron=False,
            relief="raised", width=3, padx=4
        )
        italic_btn.pack(side=tk.LEFT, padx=1)

        uline_btn = tk.Checkbutton(
            style_frame, text="S̲", font=("MS Sans Serif", 9),
            variable=self.uline_var, bg=BG_MAIN,
            command=self._schedule_preview_update, indicatoron=False,
            relief="raised", width=3, padx=4
        )
        uline_btn.pack(side=tk.LEFT, padx=1)

        # Alinhamento
        ttk.Label(row2, text="Alinhamento:").pack(side=tk.LEFT, padx=(0, 4))
        align_frame = ttk.Frame(row2)
        align_frame.pack(side=tk.LEFT)

        for label, value in [("◀ Esq", "left"), ("◆ Cen", "center"), ("▶ Dir", "right")]:
            rb = tk.Radiobutton(
                align_frame, text=label,
                variable=self.align_var, value=value,
                bg=BG_MAIN, font=FONT_UI,
                command=self._schedule_preview_update,
                indicatoron=False, relief="raised", padx=4, pady=2
            )
            rb.pack(side=tk.LEFT, padx=1)

    # ── Área de texto ────────────────────────────────────────────────────────

    def _build_textarea(self, parent):
        txt_frame = ttk.LabelFrame(parent, text="Texto", padding=5)
        txt_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 6))

        # ScrollBar
        sb = ttk.Scrollbar(txt_frame)
        sb.pack(side=tk.RIGHT, fill=tk.Y)

        self.text_box = tk.Text(
            txt_frame,
            font=("Consolas", 11),
            bg=BG_WHITE,
            fg="#000000",
            relief="sunken",
            borderwidth=2,
            wrap="word",
            undo=True,
            yscrollcommand=sb.set,
        )
        self.text_box.pack(fill=tk.BOTH, expand=True)
        sb.config(command=self.text_box.yview)

        # Atualizar preview a cada tecla (com debounce)
        self._preview_job = None
        self.text_box.bind("<KeyRelease>", lambda e: self._schedule_preview_update())

        # Texto inicial
        self.text_box.insert("1.0", "Digite seu texto aqui...")
        self.text_box.tag_add("sel", "1.0", "end")

    # ── Botões de ação ────────────────────────────────────────────────────────

    def _build_action_buttons(self, parent):
        btn_frame = ttk.Frame(parent, style="Card.TFrame")
        btn_frame.pack(fill=tk.X)

        ttk.Button(
            btn_frame, text="🖨️  Imprimir",
            command=self.do_print,
            style="Primary.TButton"
        ).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))

        ttk.Button(
            btn_frame, text="💾  Salvar Imagem",
            command=self.do_save,
            style="Secondary.TButton"
        ).pack(side=tk.LEFT, fill=tk.X, expand=True)

    # ── Preview ──────────────────────────────────────────────────────────────

    def _build_preview(self, parent):
        preview_lf = ttk.LabelFrame(parent, text="Preview da Impressão (58mm)", padding=5)
        preview_lf.pack(fill=tk.BOTH, expand=True)

        border = tk.Frame(preview_lf, bg="#b0b0b0", relief="sunken", borderwidth=2)
        border.pack(fill=tk.BOTH, expand=True, padx=3, pady=3)

        # Canvas com scroll vertical
        self.preview_canvas = tk.Canvas(
            border,
            bg=BG_WHITE,
            highlightthickness=0,
        )
        vsb = ttk.Scrollbar(border, orient="vertical", command=self.preview_canvas.yview)
        self.preview_canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self.preview_canvas.pack(fill=tk.BOTH, expand=True)

        # Indicador de início de impressão
        self.preview_canvas.bind("<Configure>", self._on_preview_configure)

        # Rótulo informativo abaixo
        self.preview_info = tk.Label(
            preview_lf, text="", font=FONT_UI,
            bg=BG_MAIN, fg="#555555"
        )
        self.preview_info.pack()

    # ── Lógica interna ────────────────────────────────────────────────────────

    def _reload_font_combo(self):
        """Reconstrói a lista de fontes no combobox com recentes no topo."""
        font_list = build_font_list(self.all_fonts, self.font_history)
        self.font_combo["values"] = font_list

    def _on_font_selected(self, event=None):
        raw = self.font_var.get()
        # Ignorar separador
        if raw.startswith("─"):
            return
        clean = clean_font_name(raw)
        self.font_var.set(clean)
        # Atualizar histórico
        self.font_history = add_to_font_history(clean, self.font_history)
        save_font_history(self.font_history)
        self._reload_font_combo()
        self._schedule_preview_update()

    def _schedule_preview_update(self, _event=None):
        """Debounce: atualiza preview 300 ms após a última alteração."""
        if self._preview_job:
            self.win.after_cancel(self._preview_job)
        self._preview_job = self.win.after(300, self._update_preview)

    def _on_preview_configure(self, event=None):
        self._update_preview()

    def _get_current_text(self) -> str:
        return self.text_box.get("1.0", "end-1c")

    def _render(self) -> Image.Image | None:
        text = self._get_current_text().strip()
        if not text:
            return None
        try:
            return render_text_to_image(
                text=text,
                font_family=self.font_var.get(),
                font_size=max(6, self.size_var.get()),
                bold=self.bold_var.get(),
                italic=self.italic_var.get(),
                underline=self.uline_var.get(),
                align=self.align_var.get(),
                paper_width=PAPER_WIDTH_PX,
                padding=PADDING_PX,
            )
        except Exception as e:
            print(f"[TextEditor] Erro ao renderizar: {e}")
            return None

    def _update_preview(self):
        self.rendered_image = self._render()
        canvas = self.preview_canvas
        canvas.delete("all")

        cw = canvas.winfo_width()
        if cw <= 1:
            cw = 400

        if self.rendered_image is None:
            canvas.create_text(
                cw // 2, 60,
                text="(texto vazio)",
                fill="#aaaaaa",
                font=FONT_UI,
            )
            self.preview_info.config(text="")
            return

        # Escalar para caber no canvas mantendo proporção
        img = self.rendered_image.copy()
        scale = (cw - 20) / PAPER_WIDTH_PX
        new_w = int(PAPER_WIDTH_PX * scale)
        new_h = int(img.height * scale)
        img = img.resize((new_w, new_h), Image.Resampling.LANCZOS)

        self._preview_photo = ImageTk.PhotoImage(img)

        x0 = (cw - new_w) // 2
        y0 = 28

        canvas.create_image(x0, y0, anchor="nw", image=self._preview_photo)

        # Linha vermelha de início de impressão
        canvas.create_line(10, y0, cw - 10, y0, fill="red", width=2, dash=(5, 5))
        canvas.create_text(
            cw // 2, y0 - 14,
            text="INÍCIO DA IMPRESSÃO (Y=0)",
            fill="red", font=("MS Sans Serif", 7, "bold"),
        )

        # Atualizar região de scroll
        canvas.configure(scrollregion=(0, 0, cw, y0 + new_h + 20))

        # Info
        h_mm = self.rendered_image.height / 8
        self.preview_info.config(
            text=f"{PAPER_WIDTH_PX}×{self.rendered_image.height}px  |  58×{h_mm:.1f}mm"
        )

    # ── Ações ─────────────────────────────────────────────────────────────────

    def do_print(self):
        """Renderiza e envia para a impressora selecionada."""
        text = self._get_current_text().strip()
        if not text:
            messagebox.showwarning("Aviso", "Digite algum texto antes de imprimir!", parent=self.win)
            return

        img = self._render()
        if img is None:
            messagebox.showerror("Erro", "Não foi possível renderizar o texto.", parent=self.win)
            return

        printer_name = self.selected_printer_var.get()
        if printer_name and printer_name != "Nenhuma impressora encontrada":
            self.printer_handler.set_printer(printer_name)

        try:
            mono = img.convert("1")  # Monocromático para térmica
            success = self.printer_handler.print_image(mono)
            if success:
                messagebox.showinfo("Sucesso", "Texto enviado para a impressora!", parent=self.win)
            else:
                messagebox.showerror("Erro", "Falha ao enviar para a impressora.", parent=self.win)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao imprimir:\n{e}", parent=self.win)

    def do_save(self):
        """Salva a imagem renderizada em arquivo."""
        img = self._render()
        if img is None:
            messagebox.showwarning("Aviso", "Digite algum texto antes de salvar!", parent=self.win)
            return

        path = filedialog.asksaveasfilename(
            parent=self.win,
            title="Salvar imagem de texto",
            defaultextension=".png",
            filetypes=[
                ("PNG", "*.png"),
                ("JPEG", "*.jpg"),
                ("BMP", "*.bmp"),
                ("Todos os arquivos", "*.*"),
            ],
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
            messagebox.showinfo("Salvo", f"Imagem salva em:\n{path}", parent=self.win)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao salvar:\n{e}", parent=self.win)


# ── Função de abertura ────────────────────────────────────────────────────────

def open_text_editor(parent, printer_handler, image_processor, selected_printer_var):
    """Abre a janela do editor de texto (singleton por parent)."""
    editor = TextEditorWindow(parent, printer_handler, image_processor, selected_printer_var)
    editor.win.focus_set()
    return editor
