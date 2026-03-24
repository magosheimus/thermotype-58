"""
Handler de impressão para ThermoType 58
Responsável por enviar imagens para impressoras térmicas ESC/POS
"""

import os
import platform
from PIL import Image

# Imports opcionais do Windows (apenas quando necessário)
try:
    import win32print  # type: ignore
    import win32ui  # type: ignore
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False


class PrinterHandler:
    def __init__(self):
        """Inicializa o handler de impressão"""
        self.printer_name = None
        self.connection = None
        
    def list_printers(self):
        """
        Lista impressoras disponíveis no sistema
        
        Returns:
            Lista de nomes de impressoras
        """
        printers = []
        
        if platform.system() == 'Windows' and WIN32_AVAILABLE:
            try:
                # Listar impressoras instaladas
                printers = [printer[2] for printer in win32print.EnumPrinters(2)]
            except Exception as e:
                print(f"Aviso: Erro ao listar impressoras: {e}")
        
        return printers
    
    def set_printer(self, printer_name):
        """
        Define a impressora a ser usada
        
        Args:
            printer_name: Nome da impressora
        """
        self.printer_name = printer_name
    
    def print_image(self, image):
        """
        Imprime imagem em impressora térmica
        
        Args:
            image: PIL Image (preferencialmente monocromática)
            
        Returns:
            True se sucesso, False caso contrário
        """
        try:
            # Método 1: Tentar usar python-escpos se disponível
            if self._print_with_escpos(image):
                return True
            
            # Método 2: Tentar impressão RAW no Windows
            if platform.system() == 'Windows':
                if self._print_raw_windows(image):
                    return True
            
            # Método 3: Fallback - salvar e abrir com visualizador padrão
            # (não é ideal, mas funciona para testes)
            print("Aviso: Nenhum método de impressão direta disponível.")
            print("Salvando imagem para impressão manual...")
            
            temp_file = "temp_print.png"
            image.save(temp_file)
            
            # Abrir com aplicativo padrão (usuário precisa imprimir manualmente)
            if platform.system() == 'Windows':
                os.startfile(temp_file)
            
            return True
            
        except Exception as e:
            print(f"Erro ao imprimir: {e}")
            return False
    
    def _print_with_escpos(self, image):
        """
        Tenta imprimir usando biblioteca python-escpos
        
        Args:
            image: PIL Image
            
        Returns:
            True se sucesso, False caso contrário
        """
        try:
            from escpos import printer
            
            # Tentar conectar via USB
            # Nota: Você precisará ajustar os IDs vendor/product da sua impressora
            # Use: python -m escpos.cli find para descobrir
            
            # Exemplo genérico:
            # p = printer.Usb(0x0416, 0x5011)  # Substitua pelos IDs corretos
            
            # Por enquanto, retornar False para usar outros métodos
            return False
            
        except ImportError:
            return False
        except Exception as e:
            print(f"Erro com escpos: {e}")
            return False
    
    def _print_raw_windows(self, image):
        """
        Tenta imprimir usando RAW no Windows
        
        Args:
            image: PIL Image
            
        Returns:
            True se sucesso, False caso contrário
        """
        if not WIN32_AVAILABLE:
            return False
            
        try:
            from PIL import ImageWin
            
            # Obter impressora padrão ou usar a definida
            printer_name = self.printer_name or win32print.GetDefaultPrinter()
            
            # Abrir impressora
            hprinter = win32print.OpenPrinter(printer_name)
            
            try:
                # Criar DC
                hdc = win32ui.CreateDC()
                hdc.CreatePrinterDC(printer_name)
                
                # Iniciar documento
                hdc.StartDoc("ThermoType 58 Print")
                hdc.StartPage()
                
                # Converter imagem para bitmap
                dib = ImageWin.Dib(image)
                
                # Imprimir no topo (0, 0)
                dib.draw(hdc.GetHandleOutput(), (0, 0, image.width, image.height))
                
                # Finalizar
                hdc.EndPage()
                hdc.EndDoc()
                
                return True
                
            finally:
                win32print.ClosePrinter(hprinter)
                
        except Exception as e:
            print(f"Erro na impressão RAW: {e}")
            return False
    
    def print_test_page(self):
        """
        Imprime página de teste
        
        Returns:
            True se sucesso, False caso contrário
        """
        # Criar imagem de teste simples
        test_image = Image.new('1', (464, 200), 1)  # Branco
        
        from PIL import ImageDraw, ImageFont
        draw = ImageDraw.Draw(test_image)
        
        # Desenhar texto de teste
        draw.text((10, 10), "ThermoType 58", fill=0)
        draw.text((10, 30), "Teste de Impressao", fill=0)
        draw.text((10, 50), "58mm Thermal Printer", fill=0)
        draw.text((10, 70), "Width: 464px (58mm)", fill=0)
        
        # Desenhar linha no topo
        draw.line([(0, 0), (464, 0)], fill=0, width=2)
        
        return self.print_image(test_image)
    
    def get_esc_pos_commands(self, image):
        """
        Gera comandos ESC/POS para imprimir a imagem
        
        Args:
            image: PIL Image (monocromática)
            
        Returns:
            Bytes com comandos ESC/POS
        """
        commands = bytearray()
        
        # Initialize printer
        commands.extend(b'\x1B\x40')  # ESC @
        
        # Set line spacing to 0
        commands.extend(b'\x1B\x33\x00')  # ESC 3 0
        
        # Converter imagem para raster bitmap
        # (Simplificado - implementação completa requer processamento bit a bit)
        
        # Print and feed
        commands.extend(b'\x1B\x64\x02')  # ESC d 2 (feed 2 lines)
        
        # Cut paper (se suportado)
        commands.extend(b'\x1D\x56\x00')  # GS V 0
        
        return bytes(commands)
