# 🖨️ ThermoType 58

App desktop para impressoras térmicas 58mm: editor de texto com fontes do sistema, impressão direta e correção automática de margem superior.

![Versão](https://img.shields.io/badge/versão-1.2.0-blue)
![Python](https://img.shields.io/badge/python-3.8+-green)
![Licença](https://img.shields.io/badge/licença-MIT-orange)
![Windows](https://img.shields.io/badge/plataforma-Windows-0078d7)

## 🎯 Problema

Impressoras térmicas frequentemente centralizam o conteúdo verticalmente, criando grandes espaços em branco no início da impressão e desperdiçando papel. Este aplicativo detecta e remove automaticamente essas margens, forçando a impressão a começar no topo absoluto do papel.

### Objetivo:
- Eficiência e redução de desperdício de papel
- Simplicidade e controle direto
- Interface leve e funcional

---

## ✨ Funcionalidades

### 🖼️ Impressão de Imagens
- **Drag & Drop**: Arraste imagens diretamente para o preview
- **Auto Top Fix**: Remove automaticamente a margem branca superior
- **Ajuste de Largura Automático**: Redimensiona para 58mm (384px)
- **Offset Manual**: Controle fino da posição vertical (em mm)
- **Múltiplas Cópias**: Imprima várias cópias de uma vez
- **Salvar como Imagem**: Exporte em PNG, JPG ou BMP

### ✏️ Editor de Texto
- **Editor integrado**: Escreva e imprima diretamente no app
- **Fontes do Sistema**: Todas as fontes instaladas, com recentes no topo
- **Formatação completa**: Negrito, itálico, sublinhado, tamanho e alinhamento
- **Preview em Tempo Real**: Visualize exatamente como ficará a impressão
- **Formatação por seleção**: Aplique estilos diferentes em partes do texto

### 📋 Templates
- **Salvar Template**: Salve texto + formatação com um nome para reutilizar
- **Carregar Template**: Restaura texto e toda a formatação com um clique
- **Gerenciar**: Liste, edite e exclua templates diretamente na interface

### 🕓 Histórico & Undo/Redo
- **Histórico de imagens**: Acesso rápido via miniaturas às últimas imagens usadas
- **Último texto no histórico**: O último texto impresso fica salvo como miniatura
- **Undo/Redo de imagem**: Desfaça e refaça carregamentos e ajustes (`Ctrl+Z` / `Ctrl+Y`)

### ⚙️ Configurações Persistentes
- Impressora selecionada, fonte, tamanho, formatação, offset, cópias e Auto Top Fix são **salvos automaticamente** entre sessões

### 🖥️ Interface
- **Tema Vintage Windows 95/98**: Design retrô nostálgico
- **Ícone de impressora**: Identificação visual em todas as janelas e na barra de tarefas
- **Seleção de Impressora**: Escolha qual impressora usar, com atualização em tempo real

---

## 📋 Requisitos

- Windows 10+
- Python 3.8+
- Impressora térmica 58mm (ESC/POS compatível)
- Driver da impressora instalado

---

## 🚀 Instalação e Uso

### 1. Clone o repositório

```bash
git clone https://github.com/magosheimus/thermotype-58.git
cd thermotype-58
```

### 2. Crie o ambiente virtual e instale as dependências

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

### 3. Execute o aplicativo

```bash
python main.py
```

### 4. (Opcional) Criar executável standalone

```bash
python build_exe.py
```

O executável será criado em `dist/ThermoType 58.exe` — sem precisar de Python instalado.

---

## 📁 Estrutura do Projeto

```
ThermoType 58/
├── main.py                 # Aplicativo principal + UI
├── text_editor.py          # Editor de texto com formatação
├── image_processor.py      # Processamento e ajuste de imagens
├── printer_handler.py      # Interface com a impressora
├── config.py               # Configurações globais
├── build_exe.py            # Script para gerar o .exe
├── requirements.txt        # Dependências Python
├── printer.ico             # Ícone do aplicativo
└── README.md               # Este arquivo
```

**Arquivos gerados automaticamente (não versionados):**
- `editor_settings.json` — configurações persistentes do usuário
- `templates.json` — templates de texto salvos
- `history.json` — histórico de imagens
- `font_history.json` — histórico de fontes usadas
- `_last_text_preview.png` — preview do último texto impresso

---

## 🛠️ Tecnologias Utilizadas

- **Python 3.8+**: Linguagem de programação
- **Tkinter + TkinterDnD2**: Interface gráfica com drag & drop
- **Pillow (PIL)**: Processamento de imagens
- **NumPy**: Análise de pixels para Auto Top Fix
- **pywin32**: Comunicação direta com impressoras Windows
- **PyInstaller**: Geração do executável standalone

---

## 📝 Licença

Este projeto está sob a licença MIT. Veja o arquivo [LICENSE](LICENSE) para mais detalhes.

---

## 🌱 Autor

Projeto independente desenvolvido a partir de uma necessidade prática no uso diário de impressoras térmicas.

---

**⭐ Se este projeto foi útil para você, considere dar uma estrela!**
