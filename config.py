"""
ThermoType 58 - Configurações
"""

# Configurações da impressora térmica
PRINTER_CONFIG = {
    'paper_width_mm': 58,
    'dpi': 203,
    'pixels_per_mm': 8,
    'paper_width_px': 384,  # Reduzido de 464 para compensar margens da impressora
    'max_paper_height_mm': 2000,  # Altura máxima do rolo
}

# Configurações de processamento de imagem
IMAGE_CONFIG = {
    'white_threshold': 250,  # Pixels acima desse valor são considerados brancos
    'dark_pixel_ratio': 0.01,  # 1% de pixels escuros para detectar conteúdo
    'dither_method': 'floyd-steinberg',  # Método de dithering
    'default_mode': 'threshold',  # 'threshold' ou 'dither'
}

# Configurações da interface
UI_CONFIG = {
    'window_width': 800,
    'window_height': 700,
    'preview_max_height': 400,
    'canvas_width': 464,
}

# Comandos ESC/POS
ESCPOS_COMMANDS = {
    'initialize': b'\x1B\x40',  # ESC @
    'line_spacing_0': b'\x1B\x33\x00',  # ESC 3 0
    'align_left': b'\x1B\x61\x00',  # ESC a 0
    'feed_2': b'\x1B\x64\x02',  # ESC d 2
    'cut': b'\x1D\x56\x00',  # GS V 0
}
