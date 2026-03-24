# 🖨️ ThermoType 58

App desktop para impressoras térmicas 58mm: editor de texto, impressão de imagens e correção automática de margem superior.

![Versão](https://img.shields.io/badge/versão-1.2.0-blue)
![Python](https://img.shields.io/badge/python-3.8+-green)
![Licença](https://img.shields.io/badge/licença-MIT-orange)
![Windows](https://img.shields.io/badge/plataforma-Windows-0078d7)

---

## 🎯 Por que existe

Impressoras térmicas frequentemente centralizam o conteúdo verticalmente, desperdiçando papel com espaço em branco no início. O ThermoType 58 detecta e remove essa margem automaticamente, forçando a impressão a começar no topo do papel.

---

## ✨ Funcionalidades

### 🖼️ Impressão de Imagens
- **Drag & Drop**: Arraste imagens diretamente para o preview
- **Auto Top Fix**: Remove a margem branca superior automaticamente
- **Redimensionamento automático**: Ajusta a imagem para 384px (58mm)
- **Offset manual**: Controle da posição vertical em milímetros
- **Múltiplas cópias**: Defina quantas cópias imprimir
- **Exportar imagem**: Salve o resultado em PNG, JPG ou BMP

### ✏️ Editor de Texto
- **Editor integrado**: Escreva e imprima sem sair do app
- **Fontes do sistema**: Todas as fontes instaladas disponíveis, com recentes no topo
- **Formatação completa**: Negrito, itálico, sublinhado, tamanho e alinhamento
- **Preview em tempo real**: Veja exatamente como ficará antes de imprimir
- **Formatação por seleção**: Estilos diferentes em partes distintas do texto

### 📋 Templates
- Salve texto com toda a formatação e dê um nome
- Carregue qualquer template com um clique
- Gerencie (liste, edite, exclua) direto na interface

### 🕓 Histórico & Undo/Redo
- Miniaturas das últimas imagens usadas para acesso rápido
- Último texto impresso salvo como miniatura no histórico
- Desfazer/refazer carregamentos e ajustes com `Ctrl+Z` / `Ctrl+Y`

### ⚙️ Configurações persistentes
Impressora, fonte, tamanho, formatação, offset, cópias e Auto Top Fix são **salvos automaticamente** entre sessões.

### 🖥️ Interface
- Tema visual estilo Windows 95/98
- Ícone de impressora em todas as janelas e na barra de tarefas
- Seleção de impressora com atualização em tempo real

---

## 📋 Requisitos

- Windows 10 ou superior
- Python 3.8+
- Impressora térmica 58mm com driver instalado

---

## 🚀 Instalação

### 1. Clone o repositório

```bash
git clone https://github.com/magosheimus/thermotype-58.git
cd thermotype-58
```

### 2. Crie o ambiente virtual e instale as dependências

```bash
python -m venv .venv
.venv\Scripts\activate        # CMD
# ou
.venv\Scripts\Activate.ps1    # PowerShell
pip install -r requirements.txt
```

### 3. Execute

```bash
python main.py
```

### 4. (Opcional) Gerar executável `.exe`

```bash
python build_exe.py
```

O arquivo será gerado em `dist/ThermoType 58.exe` — não requer Python instalado para rodar.

---

## 📁 Estrutura do Projeto

```
ThermoType 58/
├── main.py                 # Aplicativo principal e UI
├── text_editor.py          # Editor de texto com formatação
├── image_processor.py      # Processamento e Auto Top Fix
├── printer_handler.py      # Comunicação com a impressora
├── config.py               # Configurações globais
├── build_exe.py            # Script para gerar o .exe
├── requirements.txt        # Dependências Python
└── printer.ico             # Ícone do aplicativo
```

**Gerados automaticamente (não versionados):**

| Arquivo | Conteúdo |
|---|---|
| `editor_settings.json` | Configurações do usuário |
| `templates.json` | Templates de texto salvos |
| `history.json` | Histórico de imagens |
| `font_history.json` | Histórico de fontes usadas |
| `_last_text_preview.png` | Preview do último texto impresso |

---

## 🛠️ Tecnologias

| Biblioteca | Uso |
|---|---|
| Tkinter + TkinterDnD2 | Interface gráfica e drag & drop |
| Pillow (PIL) | Processamento de imagens |
| NumPy | Análise de pixels para Auto Top Fix |
| pywin32 | Comunicação com impressoras Windows |
| PyInstaller | Geração do executável standalone |

---

## 📝 Licença

MIT — veja o arquivo [LICENSE](LICENSE).

---

## 🌱 Autor

Desenvolvido de forma independente a partir de uma necessidade real no uso diário de impressoras térmicas.

---

**⭐ Se este projeto foi útil para você, deixe uma estrela!**
