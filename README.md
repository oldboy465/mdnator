# 🌋 MDnator v3.1 — Conversor Universal & MS Engine (Lava Edition)

[![Python](https://img.shields.io/badge/Python-3.10%2B-blue?style=for-the-badge&logo=python&logoColor=white)](https://www.python.org/)
[![PyQt6](https://img.shields.io/badge/GUI-PyQt6-green?style=for-the-badge&logo=qt&logoColor=white)](https://pypi.org/project/PyQt6/)
[![Microsoft MarkItDown](https://img.shields.io/badge/Engine-Microsoft%20MarkItDown-0078D4?style=for-the-badge&logo=microsoft&logoColor=white)](https://github.com/microsoft/markitdown)
[![LangChain](https://img.shields.io/badge/Engine-LangChain-121212?style=for-the-badge&logo=chainlink&logoColor=white)](https://www.langchain.com/)
[![License](https://img.shields.io/badge/License-MIT-orange?style=for-the-badge)](LICENSE)

O **MDnator v3.1** é uma aplicação desktop de alta performance desenvolvida em **Python** e **PyQt6**, projetada para transformar documentos, arquivos e repositórios completos de código/arquivos corporativos em arquivos **Markdown (.md)** altamente estruturados.

Projetado com a estética incandescente **Lava Edition (Glassmorphism Dark)**, o MDnator integra dois potentes motores de conversão: o **LangChain Engine** tradicional para leitura rápida de documentos e o moderno **Microsoft MarkItDown Engine** para fidelidade impecável em arquivos do pacote Office (Word, Excel, PowerPoint), PDFs e arquivos de mídia.

---

## ⚡ Principais Funcionalidades

- 📄 **Conversão Universal de Documentos:**
  - Suporte a **PDF, DOCX, PPTX, XLSX, XLS, CSV, HTML, SAV (SPSS)** e dezenas de extensões de código/texto simples.
- 🚀 **Dois Motores de Processamento (Multi-Engine):**
  - **LangChain Engine:** Leitura e estruturação ágil baseada em carregadores comunitários do LangChain.
  - **Microsoft MarkItDown Engine:** Alta precisão de marcação e preservação de metadados para documentos Office e PDFs complexos.
- 🌳 **Scanner Completo de Projetos & Código:**
  - Varredura recursiva de diretórios com geração automática da árvore de arquivos e inclusão do conteúdo de código-fonte formatado em blocos Markdown com *syntax highlighting*.
- 🎛️ **Seleção Granular de Arquivos (Checkboxes Dinâmicos):**
  - Renderização assíncrona da árvore do projeto com opções de marcar/desmarcar arquivos individualmente antes de processar.
- 🎨 **Interface Incandescente "Lava Edition":**
  - Design futurista com paleta *Obsidian & Magma*, suporte nativo ao *Qt Fusion*, sombras dinâmicas, terminal de logs integrado e barra de progresso em tempo real.
- ⚡ **Thread-Safety Executivo:**
  - Processamento assíncrono em background sem congelamento da interface gráfica.

---

## 🛠️ Tecnologias e Dependências

- **Linguagem:** Python 3.10+
- **GUI:** PyQt6
- **Motores de Conversão:**
  - `markitdown` (Microsoft)
  - `langchain-community` & `langchain-core`
- **Manipulação de Dados:** `pandas`, `pyreadstat` (arquivos `.sav` do SPSS)
- **Leitores de Documentos:** `pypdf`, `docx2txt`, `unstructured`

---

## 📂 Estrutura de Arquivos Suportados

| Categoria | Extensões Suportadas |
| :--- | :--- |
| **Documentos Office** | `.docx`, `.pptx`, `.xlsx`, `.xls` |
| **Documentos Gerais** | `.pdf`, `.csv`, `.sav` (SPSS), `.html`, `.htm` |
| **Linguagens de Programação** | `.py`, `.js`, `.ts`, `.php`, `.java`, `.c`, `.cpp`, `.h`, `.cs`, `.sql` |
| **Web & Estilos** | `.html`, `.htm`, `.css`, `.json`, `.xml` |
| **Scripts & Configuração** | `.sh`, `.bat`, `.cmd`, `.md`, `.txt` |

---

## 🚀 Como Executar o Projeto

### 1️⃣ Clonar o Repositório
```bash
git clone [https://github.com/oldboy465/mdnator.git](https://github.com/oldboy465/mdnator.git)
cd mdnator