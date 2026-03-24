"""
Processador de imagem para TopStart Thermal
Responsável por ajustar, recortar e preparar imagens para impressão térmica
"""

from PIL import Image, ImageOps
import numpy as np


class ImageProcessor:
    def __init__(self, target_width_px=464, pixels_per_mm=8):
        """
        Inicializa o processador de imagem
        
        Args:
            target_width_px: Largura alvo em pixels (464 para 58mm)
            pixels_per_mm: Pixels por milímetro (8 para 203 DPI)
        """
        self.target_width_px = target_width_px
        self.pixels_per_mm = pixels_per_mm
    
    def resize_to_width(self, image):
        """
        Redimensiona proporcionalmente para largura máxima de 464px, centralizando sempre.
        Args:
            image: PIL Image
        Returns:
            PIL Image centralizada
        """
        if image.mode not in ('RGB', 'L'):
            image = image.convert('RGB')
        
        print(f"[DEBUG] Imagem original: {image.width}x{image.height}px")
        print(f"[DEBUG] Largura alvo: {self.target_width_px}px")
        
        # Se a imagem for mais larga, reduz proporcionalmente
        if image.width > self.target_width_px:
            scale = self.target_width_px / image.width
            new_width = self.target_width_px
            new_height = int(image.height * scale)
            print(f"[DEBUG] Reduzindo para: {new_width}x{new_height}px")
            resized = image.resize((new_width, new_height), Image.Resampling.LANCZOS)
        # Se for menor, mantém tamanho e centraliza
        elif image.width < self.target_width_px:
            new_height = image.height
            new_img = Image.new('RGB', (self.target_width_px, new_height), 'white')
            offset_x = int((self.target_width_px - image.width) / 2)
            print(f"[DEBUG] Centralizando com offset X: {offset_x}px")
            new_img.paste(image, (offset_x, 0))  # Centraliza
            resized = new_img
        else:
            print(f"[DEBUG] Largura já está correta")
            resized = image
        
        print(f"[DEBUG] Imagem final: {resized.width}x{resized.height}px")
        return resized
    
    def remove_top_margin(self, image, threshold=250):
        """
        Remove margem branca superior da imagem
        
        Args:
            image: PIL Image
            threshold: Valor de luminosidade para considerar como branco (0-255)
            
        Returns:
            PIL Image com margem superior removida
        """
        # Converter para grayscale para análise
        if image.mode != 'L':
            gray_image = image.convert('L')
        else:
            gray_image = image
        
        # Converter para array numpy
        img_array = np.array(gray_image)
        
        # Encontrar primeira linha com pixels escuros
        top_line = 0
        for y in range(img_array.shape[0]):
            # Contar pixels escuros na linha
            dark_pixels = np.sum(img_array[y] < threshold)
            
            # Se encontrar pixels escuros suficientes (>1% da largura)
            if dark_pixels > (img_array.shape[1] * 0.01):
                top_line = y
                break
        
        # Se encontrou conteúdo, recortar
        if top_line > 0:
            # Recortar a imagem original (não a grayscale)
            cropped = image.crop((0, top_line, image.width, image.height))
            return cropped
        
        return image
    
    def apply_offset(self, image, offset_mm):
        """
        Aplica offset vertical manual
        
        Args:
            image: PIL Image
            offset_mm: Offset em milímetros (positivo = para baixo, negativo = para cima)
            
        Returns:
            PIL Image com offset aplicado
        """
        offset_px = int(offset_mm * self.pixels_per_mm)
        
        if offset_px == 0:
            return image
        
        # Criar nova imagem com espaço para offset
        if offset_px > 0:
            # Adicionar espaço branco no topo
            new_height = image.height + offset_px
            new_image = Image.new('RGB', (image.width, new_height), 'white')
            new_image.paste(image, (0, offset_px))
            return new_image
        else:
            # Remover pixels do topo
            offset_px = abs(offset_px)
            if offset_px < image.height:
                return image.crop((0, offset_px, image.width, image.height))
            else:
                return image
    
    def convert_to_monochrome(self, image, method='threshold'):
        """
        Converte imagem para monocromático para impressão térmica
        
        Args:
            image: PIL Image
            method: 'threshold' ou 'dither'
            
        Returns:
            PIL Image em modo '1' (preto e branco puro)
        """
        # Converter para grayscale primeiro
        if image.mode != 'L':
            gray = image.convert('L')
        else:
            gray = image
        
        if method == 'threshold':
            # Threshold simples em 50%
            mono = gray.point(lambda x: 0 if x < 128 else 255, '1')
        elif method == 'dither':
            # Dithering (Floyd-Steinberg)
            mono = gray.convert('1')
        else:
            mono = gray.convert('1')
        
        return mono
    
    def detect_content_height(self, image, threshold=250):
        """
        Detecta a altura real do conteúdo (ignorando margens brancas)
        
        Args:
            image: PIL Image
            threshold: Valor de luminosidade para considerar como branco
            
        Returns:
            Altura do conteúdo em pixels
        """
        # Converter para grayscale
        if image.mode != 'L':
            gray_image = image.convert('L')
        else:
            gray_image = image
        
        img_array = np.array(gray_image)
        
        # Encontrar última linha com conteúdo
        bottom_line = image.height
        for y in range(img_array.shape[0] - 1, -1, -1):
            dark_pixels = np.sum(img_array[y] < threshold)
            if dark_pixels > (img_array.shape[1] * 0.01):
                bottom_line = y + 1
                break
        
        # Encontrar primeira linha com conteúdo
        top_line = 0
        for y in range(img_array.shape[0]):
            dark_pixels = np.sum(img_array[y] < threshold)
            if dark_pixels > (img_array.shape[1] * 0.01):
                top_line = y
                break
        
        return bottom_line - top_line
    
    def auto_crop_content(self, image):
        """
        Recorta automaticamente para manter apenas o conteúdo
        
        Args:
            image: PIL Image
            
        Returns:
            PIL Image recortada
        """
        # Remover margem superior
        image = self.remove_top_margin(image)
        
        # Detectar altura do conteúdo
        content_height = self.detect_content_height(image)
        
        # Recortar para altura do conteúdo
        if content_height < image.height:
            image = image.crop((0, 0, image.width, content_height))
        
        return image
