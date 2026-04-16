import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
import time
import threading
from datetime import datetime
from typing import List, Tuple, Optional
import hashlib

# --- IMPORTAÇÕES DO LANGCHAIN (MÓDULO ORIGINAL) ---
try:
    from langchain_community.document_loaders import (
        PyPDFLoader,
        Docx2txtLoader,
        TextLoader,
        BSHTMLLoader,
        CSVLoader
    )
    from langchain_community.document_loaders.excel import UnstructuredExcelLoader as PandasExcelLoader
    from langchain_core.documents import Document
except ImportError:
    pass # Tratado no main

import pandas as pd
import pyreadstat

# --- IMPORTAÇÃO DO NOVO MÓDULO (MARKITDOWN MICROSOFT) ---
try:
    from markitdown import MarkItDown
    MARKITDOWN_AVAILABLE = True
except ImportError:
    MARKITDOWN_AVAILABLE = False


class MDConverter:
    """
    O CÉREBRO ORIGINAL (LangChain)
    """
    def __init__(self, log_callback=None):
        self.log = log_callback if log_callback else lambda msg: print(f"[MDConverter] {msg}")

    def load_document(self, file_path: str) -> Tuple[List[Document], dict]:
        self.log(f"Analisando (LangChain): {os.path.basename(file_path)}")
        _, file_extension = os.path.splitext(file_path)
        file_extension = file_extension.lower()
        documents = []
        metadata = {}

        try:
            if file_extension == '.pdf':
                loader = PyPDFLoader(file_path)
                documents = loader.load_and_split()
            elif file_extension == '.docx':
                loader = Docx2txtLoader(file_path)
                documents = loader.load()
            elif file_extension in ['.txt', '.md', '.py', '.php', '.css', '.js', '.java', '.sql', '.html', '.htm', '.json', '.xml', '.cmd', '.bat', '.c', '.cs', '.cpp', '.h']:
                loader = TextLoader(file_path, encoding='utf-8')
                documents = loader.load()
            elif file_extension == '.csv':
                loader = CSVLoader(file_path, encoding='utf-8')
                documents = loader.load()
            elif file_extension in ['.xls', '.xlsx']:
                loader = PandasExcelLoader(file_path, sheet_name=None)
                documents = loader.load()
            elif file_extension == '.html' or file_extension == '.htm':
                loader = BSHTMLLoader(file_path, open_encoding='utf-8')
                documents = loader.load()
            elif file_extension == '.sav':
                df, meta = pyreadstat.read_sav(file_path)
                md_content = df.to_markdown(index=False)
                doc = Document(page_content=md_content, metadata={"source": file_path})
                documents = [doc]
            else:
                raise ValueError(f"Formato {file_extension} não suportado pelo motor padrão.")

            if documents:
                metadata = documents[0].metadata
            return documents, metadata
        except Exception as e:
            self.log(f"ERRO LangChain: {e}")
            raise

    def format_to_markdown(self, documents: List[Document]) -> str:
        if not documents: return "# ERRO: Nenhum conteúdo."
        output_md = [f"# 🧠 Conversão Padrão (LangChain)\n\n"]
        for doc in documents:
            output_md.append(doc.page_content.strip())
            output_md.append("\n\n---\n\n")
        return "".join(output_md)


class MSEngine:
    """
    NOVO CÉREBRO: Microsoft MarkItDown
    Focado em alta fidelidade para Office e PDF.
    """
    def __init__(self, log_callback=None):
        self.log = log_callback if log_callback else lambda msg: print(f"[MSEngine] {msg}")
        if MARKITDOWN_AVAILABLE:
            self.md = MarkItDown()
        else:
            self.md = None

    def convert_file(self, file_path: str) -> str:
        if not self.md:
            raise ImportError("Biblioteca 'markitdown' não instalada. Execute: pip install markitdown")
        
        self.log(f"Processando com Microsoft MarkItDown: {os.path.basename(file_path)}")
        try:
            result = self.md.convert(file_path)
            
            header = f"# 📷 Microsoft MarkItDown Report\n"
            header += f"**Arquivo:** `{os.path.basename(file_path)}`\n"
            header += f"**Data:** {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n\n---\n\n"
            
            return header + result.text_content
        except Exception as e:
            self.log(f"ERRO MarkItDown: {e}")
            raise

    def scan_project(self, root_dir: str, selected_files: Optional[List[str]] = None) -> str:
        """Varre projeto usando MarkItDown para arquivos suportados, filtrando os selecionados."""
        if not self.md:
            raise ImportError("Biblioteca 'markitdown' não instalada.")
            
        self.log(f"Iniciando varredura avançada (MS) em: {root_dir}")
        output = [f"# 🌳 Projeto Completo (Engine Microsoft)\n**Raiz:** `{root_dir}`\n\n"]
        
        supported_exts = {'.pptx', '.docx', '.xlsx', '.pdf', '.jpg', '.png', '.html', '.csv', '.json', '.xml', '.txt', '.md', '.py', '.js', '.css', '.c', '.cpp', '.cs', '.bat', '.cmd'}
        
        for dirpath, _, filenames in os.walk(root_dir):
            if any(x in dirpath for x in ['__pycache__', '.git', 'node_modules', 'venv']):
                continue
                
            for filename in filenames:
                _, ext = os.path.splitext(filename)
                if ext.lower() in supported_exts:
                    full_path = os.path.join(dirpath, filename)
                    
                    # Filtro de Checkbox: Só processa se estiver na lista de selecionados
                    if selected_files is not None and full_path not in selected_files:
                        continue

                    rel_path = os.path.relpath(full_path, root_dir)
                    
                    self.log(f"Convertendo: {rel_path}")
                    
                    # Tenta contar as linhas se for arquivo de texto
                    try:
                        with open(full_path, 'r', encoding='utf-8') as f:
                            lines = len(f.readlines())
                            linha_info = f" ({lines} linhas)"
                    except:
                        linha_info = " (Arquivo Binário/Não textual)"

                    output.append(f"\n## 📄 Arquivo: `{rel_path}`{linha_info}\n\n")
                    try:
                        res = self.md.convert(full_path)
                        output.append(res.text_content)
                    except Exception as e:
                        output.append(f"> ⚠️ Erro ao converter: {e}")
                    output.append("\n\n---\n\n")
                    
        return "".join(output)


class ProjectScanner:
    """CÉREBRO ORIGINAL DE CÓDIGO (Atualizado com novas extensões e contagem de linhas)"""
    TEXT_EXTENSIONS = {'.py', '.js', '.css', '.html', '.md', '.txt', '.json', '.xml', '.sql', '.java', '.c', '.cpp', '.h', '.cs', '.sh', '.bat', '.php', '.cmd'}
    LANG_MAP = {'.py': 'python', '.js': 'javascript', '.css': 'css', '.html': 'html', '.md': 'markdown', '.json': 'json', '.sql': 'sql', '.php': 'php', '.c': 'c', '.cpp': 'cpp', '.cs': 'csharp', '.bat': 'bat', '.cmd': 'bat', '.sh': 'bash'}

    def __init__(self, log_callback=None):
        self.log = log_callback if log_callback else lambda msg: print(f"[Scanner] {msg}")
        self.ignored_dirs = {'__pycache__', '.git', '.vscode', 'node_modules', 'venv', 'env'}

    def scan_directory(self, root_dir: str, selected_files: Optional[List[str]] = None) -> str:
        self.log(f"Varrendo código (Simples): {root_dir}")
        structure = [f"# 🌳 Relatório de Código\n**Raiz:** `{root_dir}`\n\n## Estrutura\n```\n"]
        contents = ["\n\n## Conteúdo\n"]
        
        for dirpath, dirnames, filenames in os.walk(root_dir):
            dirnames[:] = [d for d in dirnames if d not in self.ignored_dirs]
            rel_dir = os.path.relpath(dirpath, root_dir)
            prefix = "" if rel_dir == "." else "│   " * (rel_dir.count(os.sep) + 1)
            
            if rel_dir != ".": structure.append(f"{'│   ' * rel_dir.count(os.sep)}├── 📁 {os.path.basename(dirpath)}/\n")

            for filename in filenames:
                _, ext = os.path.splitext(filename)
                if ext.lower() in self.TEXT_EXTENSIONS:
                    file_path = os.path.join(dirpath, filename)
                    
                    # Verifica se o arquivo foi marcado na interface
                    if selected_files is not None and file_path not in selected_files:
                        continue

                    rel_file = os.path.relpath(file_path, root_dir)
                    structure.append(f"{prefix}├── 📄 {filename}\n")
                    
                    try:
                        # Lê todas as linhas para conseguir fazer a contagem
                        with open(file_path, 'r', encoding='utf-8') as f: 
                            lines = f.readlines()
                            line_count = len(lines)
                            file_content = "".join(lines)
                            
                        # Insere a contagem de linhas dinamicamente no título do ficheiro
                        contents.append(f"\n### 📄 `{rel_file}` ({line_count} linhas)\n```{self.LANG_MAP.get(ext.lower(), '')}\n")
                        contents.append(file_content)
                    except: 
                        contents.append(f"\n### 📄 `{rel_file}` (Erro)\n```{self.LANG_MAP.get(ext.lower(), '')}\n")
                        contents.append("Erro de leitura ou encoding.")
                        
                    contents.append("\n```\n")
                    
        structure.append("```\n")
        return "".join(structure) + "".join(contents)


class MDnatorApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("MDnator v3.1 - Conversor Universal & MS Engine")
        self.geometry("1100x800+100+50") # Aumentei um pouco a janela base para caber a nova UI
        self.minsize(900, 750)
        
        self.style = ttk.Style()
        self.style.theme_use('clam')
        self.bg_color = "#f0f4f8"
        self.accent_color = "#2563eb"
        self.configure(bg=self.bg_color)
        
        self.style.configure('Title.TLabel', font=('Segoe UI', 11, 'bold'), background=self.bg_color)
        self.style.configure('Header.TLabelframe', background=self.bg_color)
        
        self.source_file_path = None
        self.source_project_path = None
        
        # Dicionário para guardar as variáveis de cada checkbox de ficheiro
        self.file_check_vars = {}
        
        # --- CÉREBROS ---
        self.converter = MDConverter(log_callback=self.log_message)
        self.project_scanner = ProjectScanner(log_callback=self.log_message)
        self.ms_engine = MSEngine(log_callback=self.log_message)

        self.create_header()
        self.create_menu()
        self.create_widgets()
        self.protocol("WM_DELETE_WINDOW", self.on_exit_request)

    def create_header(self):
        header = tk.Frame(self, bg=self.accent_color, height=70)
        header.pack(fill=tk.X, side=tk.TOP)
        header.pack_propagate(False)
        tk.Label(header, text="📄 MDnator v3.1", font=('Segoe UI', 18, 'bold'), bg=self.accent_color, fg='white').pack(side=tk.LEFT, padx=20)
        tk.Label(header, text="Classic Converter + Microsoft MarkItDown", font=('Segoe UI', 10), bg=self.accent_color, fg='white').pack(side=tk.LEFT, pady=15)

    def create_menu(self):
        self.menu_bar = tk.Menu(self)
        self.config(menu=self.menu_bar)
        
        # Menu Arquivo
        file_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="Arquivo", menu=file_menu)
        file_menu.add_command(label="📂 Abrir documento...", command=self.on_open_file)
        file_menu.add_command(label="❌ Sair", command=self.on_exit_request)

        # Menu Projeto
        self.project_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="Projeto", menu=self.project_menu)
        self.project_menu.add_command(label="🌳 Selecionar Pasta...", command=self.on_open_project_folder)
        self.project_menu.add_command(label="🚀 Processar (Scanner Simples)", command=self.on_process_project_start, state="disabled")

        # Menu Executar (Padrão)
        self.exec_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="Executar (Padrão)", menu=self.exec_menu)
        self.exec_menu.add_command(label="⚡ Processar Arquivo (LangChain)", command=self.on_process_file_start, state="disabled")

        # --- NOVO MENU MARKITDOWN ---
        self.ms_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="MarkItDown (MS)", menu=self.ms_menu)
        self.ms_menu.add_command(label="⚡ Converter Arquivo (MS Engine)", command=self.on_ms_file_start, state="disabled")
        self.ms_menu.add_command(label="🚀 Converter Projeto (MS Engine)", command=self.on_ms_project_start, state="disabled")

        # Menu Ajuda
        help_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="Ajuda", menu=help_menu)
        help_menu.add_command(label="ℹ️ Sobre o MDnator...", command=self.on_about)

    def create_widgets(self):
        main = tk.Frame(self, bg=self.bg_color)
        main.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Metadados
        meta_frame = tk.Frame(main, bg=self.bg_color)
        meta_frame.pack(fill=tk.X, pady=(0, 10))
        meta_frame.columnconfigure(0, weight=1); meta_frame.columnconfigure(1, weight=1)

        # Origem
        src_frame = ttk.LabelFrame(meta_frame, text="📥 Origem", padding=15, style='Header.TLabelframe')
        src_frame.grid(row=0, column=0, sticky="nsew", padx=5)
        self.lbl_src = tk.Label(src_frame, text="Nenhum item selecionado", bg=self.bg_color, anchor='w')
        self.lbl_src.pack(fill=tk.X)

        # Destino
        dest_frame = ttk.LabelFrame(meta_frame, text="📤 Status", padding=15, style='Header.TLabelframe')
        dest_frame.grid(row=0, column=1, sticky="nsew", padx=5)
        self.lbl_status = tk.Label(dest_frame, text="Aguardando...", bg=self.bg_color, anchor='w')
        self.lbl_status.pack(fill=tk.X)

        # --- NOVA ÁREA: SELEÇÃO DE FICHEIROS COM CHECKBOXES ---
        self.files_frame = ttk.LabelFrame(main, text="☑️ Ficheiros do Projeto (Selecione para incluir)", padding=10, style='Header.TLabelframe')
        self.files_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # Botões de controlo da seleção
        btn_frame = tk.Frame(self.files_frame, bg=self.bg_color)
        btn_frame.pack(fill=tk.X, pady=(0, 5))
        
        tk.Button(btn_frame, text="Marcar Todos", command=self.select_all_files, bg="#e2e8f0", relief=tk.FLAT, cursor="hand2").pack(side=tk.LEFT, padx=2)
        tk.Button(btn_frame, text="Desmarcar Todos", command=self.deselect_all_files, bg="#e2e8f0", relief=tk.FLAT, cursor="hand2").pack(side=tk.LEFT, padx=2)

        # Canvas e Scrollbar para fazer a lista rolar
        self.canvas = tk.Canvas(self.files_frame, bg="white", highlightthickness=1, highlightbackground="#ccc")
        self.scrollbar = ttk.Scrollbar(self.files_frame, orient="vertical", command=self.canvas.yview)
        self.scrollable_file_frame = tk.Frame(self.canvas, bg="white")

        self.scrollable_file_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        
        # Evento do scroll do rato para facilitar a navegação
        self.canvas.bind_all("<MouseWheel>", lambda e: self.canvas.yview_scroll(int(-1*(e.delta/120)), "units"))

        self.canvas.create_window((0, 0), window=self.scrollable_file_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        # Log
        log_frame = ttk.LabelFrame(main, text="📝 Log", padding=10, style='Header.TLabelframe')
        log_frame.pack(fill=tk.BOTH, expand=True)
        self.log_text = tk.Text(log_frame, bg="#1e293b", fg="#e2e8f0", font=('Consolas', 9), height=8)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        self.progress = ttk.Progressbar(main, orient=tk.HORIZONTAL, mode='indeterminate')
        self.progress.pack(fill=tk.X, pady=5)

    def log_message(self, msg):
        self.after(0, lambda: self.log_text.insert(tk.END, f"[{datetime.now().strftime('%H:%M:%S')}] {msg}\n"))
        self.after(0, lambda: self.log_text.see(tk.END))

    # --- SELETORES ---
    def on_open_file(self):
        path = filedialog.askopenfilename()
        if path:
            self.source_file_path = path; self.source_project_path = None
            self.lbl_src.config(text=f"ARQUIVO: {os.path.basename(path)}")
            self.update_menus(file_mode=True)
            self.clear_file_list() # Limpa a lista se abriu apenas arquivo
            self.log_message(f"Arquivo carregado: {path}")

    def on_open_project_folder(self):
        path = filedialog.askdirectory()
        if path:
            self.source_project_path = path; self.source_file_path = None
            self.lbl_src.config(text=f"PROJETO: {os.path.basename(path)}")
            self.update_menus(file_mode=False)
            self.populate_file_list(path) # Preenche a nova UI de checkboxes
            self.log_message(f"Pasta carregada e ficheiros listados: {path}")

    def update_menus(self, file_mode):
        state_file = "normal" if file_mode else "disabled"
        state_proj = "disabled" if file_mode else "normal"
        
        self.exec_menu.entryconfig(0, state=state_file)
        self.project_menu.entryconfig(1, state=state_proj)
        self.ms_menu.entryconfig(0, state=state_file)
        self.ms_menu.entryconfig(1, state=state_proj)

    # --- LÓGICA DA UI DE CHECKBOXES ---
    def clear_file_list(self):
        for widget in self.scrollable_file_frame.winfo_children():
            widget.destroy()
        self.file_check_vars.clear()

    def select_all_files(self):
        for var in self.file_check_vars.values():
            var.set(True)

    def deselect_all_files(self):
        for var in self.file_check_vars.values():
            var.set(False)

    def populate_file_list(self, root_dir):
        self.clear_file_list()
        
        # Juntamos as extensões de ambos os motores para listar tudo
        ms_exts = {'.pptx', '.docx', '.xlsx', '.pdf', '.jpg', '.png', '.html', '.csv', '.json', '.xml', '.txt', '.md'}
        all_exts = self.project_scanner.TEXT_EXTENSIONS.union(ms_exts)
        
        for dirpath, dirnames, filenames in os.walk(root_dir):
            dirnames[:] = [d for d in dirnames if d not in self.project_scanner.ignored_dirs]
            
            for filename in filenames:
                _, ext = os.path.splitext(filename)
                if ext.lower() in all_exts:
                    full_path = os.path.join(dirpath, filename)
                    rel_path = os.path.relpath(full_path, root_dir)
                    
                    var = tk.BooleanVar(value=True) # Por padrão todos vêm marcados
                    self.file_check_vars[full_path] = var
                    
                    cb = tk.Checkbutton(
                        self.scrollable_file_frame, 
                        text=rel_path, 
                        variable=var, 
                        bg="white", 
                        anchor='w',
                        font=('Segoe UI', 9)
                    )
                    cb.pack(fill=tk.X, padx=5, pady=2)
                    
        self.log_message(f"Encontrados {len(self.file_check_vars)} ficheiros suportados.")

    def get_selected_files(self) -> List[str]:
        """Retorna uma lista com os caminhos completos dos ficheiros que estão marcados na UI."""
        return [path for path, var in self.file_check_vars.items() if var.get()]

    # --- PROCESSADORES (THREADS) ---
    
    # 1. Arquivo Padrão (LangChain)
    def on_process_file_start(self):
        self.run_thread(self._process_file_lc)

    def _process_file_lc(self):
        docs, _ = self.converter.load_document(self.source_file_path)
        md = self.converter.format_to_markdown(docs)
        self.save_file(md)

    # 2. Projeto Padrão (Scanner)
    def on_process_project_start(self):
        self.run_thread(self._process_proj_scanner)

    def _process_proj_scanner(self):
        # Aqui injetamos a lista de ficheiros selecionados na UI
        selected_files = self.get_selected_files()
        md = self.project_scanner.scan_directory(self.source_project_path, selected_files)
        self.save_file(md)

    # 3. Arquivo NOVO (MarkItDown)
    def on_ms_file_start(self):
        if not MARKITDOWN_AVAILABLE:
            messagebox.showerror("Erro", "Instale a lib: pip install markitdown")
            return
        self.run_thread(self._process_file_ms)

    def _process_file_ms(self):
        md = self.ms_engine.convert_file(self.source_file_path)
        self.save_file(md, suffix="_MS")

    # 4. Projeto NOVO (MarkItDown)
    def on_ms_project_start(self):
        if not MARKITDOWN_AVAILABLE:
            messagebox.showerror("Erro", "Instale a lib: pip install markitdown")
            return
        self.run_thread(self._process_proj_ms)

    def _process_proj_ms(self):
        # Injetamos a seleção também no motor MS para sermos consistentes
        selected_files = self.get_selected_files()
        md = self.ms_engine.scan_project(self.source_project_path, selected_files)
        self.save_file(md, suffix="_ProjectMS")

    # --- UTILITÁRIOS ---
    def run_thread(self, target):
        self.progress.start(10)
        threading.Thread(target=self._wrapper, args=(target,), daemon=True).start()

    def _wrapper(self, target_func):
        try:
            target_func()
        except Exception as e:
            self.log_message(f"ERRO FATAL: {e}")
            messagebox.showerror("Erro", str(e))
        finally:
            self.after(0, self.progress.stop)

    def save_file(self, content, suffix=""):
        name = "output" + suffix + ".md"
        path = filedialog.asksaveasfilename(defaultextension=".md", initialfile=name)
        if path:
            with open(path, 'w', encoding='utf-8') as f: f.write(content)
            self.log_message(f"Salvo em: {path}")
            messagebox.showinfo("Sucesso", "Arquivo salvo!")

    def copy_to_clipboard(self, text):
        """Copia texto para a área de transferência."""
        self.clipboard_clear()
        self.clipboard_append(text)
        self.update() 
        self.log_message(f"📋 '{text}' copiado para a área de transferência!")

    def on_about(self):
        """Exibe a janela 'Sobre' com os créditos."""
        self.log_message("ℹ️ Exibindo janela 'Sobre'...")
        
        about_window = tk.Toplevel(self)
        about_window.title("Sobre o MDnator v3.1")
        about_window.geometry("550x700")
        about_window.resizable(False, False)
        about_window.configure(bg='white')
        about_window.transient(self)
        about_window.grab_set()
        
        # Cabeçalho
        header_frame = tk.Frame(about_window, bg=self.accent_color, height=100)
        header_frame.pack(fill=tk.X)
        header_frame.pack_propagate(False)
        
        title_label = tk.Label(
            header_frame,
            text="📄 MDnator v3.1",
            font=('Segoe UI', 20, 'bold'),
            bg=self.accent_color,
            fg='white'
        )
        title_label.pack(pady=20)
        
        subtitle = tk.Label(
            header_frame,
            text="Conversor Universal & Scanner de Projeto",
            font=('Segoe UI', 10),
            bg=self.accent_color,
            fg='white'
        )
        subtitle.pack()
        
        # Conteúdo
        content_frame = tk.Frame(about_window, bg='white', padx=30, pady=20)
        content_frame.pack(fill=tk.BOTH, expand=True)
        
        info_text = """
Desenvolvido por: Philipe Sampaio Lima
Em parceria com: Claude AI (Anthropic) & Gemini

Versão: 3.1 (Edição "Microsoft Engine + Filtros")
Data de lançamento: Janeiro 2026

════════════════════════════════════════

🔧 Contatos:
"""
        
        info_label = tk.Label(
            content_frame,
            text=info_text,
            font=('Consolas', 9),
            bg='white',
            fg='#334155',
            justify=tk.LEFT,
            anchor='w'
        )
        info_label.pack(fill=tk.X)
        
        # Frame para email com botão de copiar
        email_frame = tk.Frame(content_frame, bg='white')
        email_frame.pack(fill=tk.X, pady=(5, 10))
        
        email_label = tk.Label(
            email_frame,
            text="   • Email: cienciaegestao@gmail.com",
            font=('Consolas', 9),
            bg='white',
            fg='#334155',
            anchor='w'
        )
        email_label.pack(side=tk.LEFT)
        
        email_copy_btn = tk.Button(
            email_frame,
            text="📋 Copiar",
            command=lambda: self.copy_to_clipboard("cienciaegestao@gmail.com"),
            bg='#e2e8f0',
            fg='#334155',
            font=('Segoe UI', 8),
            relief=tk.FLAT,
            padx=8,
            pady=2,
            cursor='hand2'
        )
        email_copy_btn.pack(side=tk.RIGHT, padx=(10, 0))
        
        # Frame para telefone com botão de copiar
        phone_frame = tk.Frame(content_frame, bg='white')
        phone_frame.pack(fill=tk.X, pady=(0, 10))
        
        phone_label = tk.Label(
            phone_frame,
            text="   • WhatsApp: +55 98 98250-6920",
            font=('Consolas', 9),
            bg='white',
            fg='#334155',
            anchor='w'
        )
        phone_label.pack(side=tk.LEFT)
        
        phone_copy_btn = tk.Button(
            phone_frame,
            text="📋 Copiar",
            command=lambda: self.copy_to_clipboard("+55 98 98250-6920"),
            bg='#e2e8f0',
            fg='#334155',
            font=('Segoe UI', 8),
            relief=tk.FLAT,
            padx=8,
            pady=2,
            cursor='hand2'
        )
        phone_copy_btn.pack(side=tk.RIGHT, padx=(10, 0))
        
        # Restante do texto
        rest_text = """
   • GitHub: @OldBoy465

════════════════════════════════════════

🎯 Modo 1: Conversor de Arquivo (LangChain)
   • PDF, DOCX, TXT, MD
   • XLS, XLSX, CSV
   • HTML, HTM
   • SPSS (.sav)

🎯 Modo 2: Scanner de Projeto (Simples)
   • Varre pastas e subpastas (agora com Filtros UI)
   • Extrai .py, .js, .css, .sql, .html, .cs, .c, etc.
   • Gera um .md único com a árvore, código e qtde linhas.

🎯 Modo 3: Microsoft MarkItDown Engine (NOVO!)
   • Alta fidelidade para PPTX, DOCX, XLSX, PDF
   • Suporte para imagens (JPG, PNG)
   • Conversão avançada de documentos Office
   • Modo arquivo único e projeto completo com filtros

════════════════════════════════════════

© 2026 - Todos os direitos reservados
"""
        
        rest_label = tk.Label(
            content_frame,
            text=rest_text,
            font=('Consolas', 9),
            bg='white',
            fg='#334155',
            justify=tk.LEFT,
            anchor='w'
        )
        rest_label.pack(fill=tk.BOTH, expand=True)
        
        # Botão de fechar
        close_btn = tk.Button(
            about_window,
            text="Fechar",
            command=about_window.destroy,
            bg=self.accent_color,
            fg='white',
            font=('Segoe UI', 10, 'bold'),
            relief=tk.FLAT,
            padx=20,
            pady=10,
            cursor='hand2'
        )
        close_btn.pack(pady=20)
        
        # Centralizar janela
        about_window.transient(self)
        about_window.grab_set()
        self.wait_window(about_window)

    def on_exit_request(self):
        """Mensagem personalizada ao fechar o programa."""
        self.log_message("👋 Usuário tentou sair...")
        
        response = messagebox.askyesno(
            "Já vai embora?", 
            "Tem certeza que já vai, baby? 🥺\n\n"
            "O que foi? Só me usou e jogou fora.... kkk\n\n"
            "Deseja realmente sair do MDnator?",
            icon='question'
        )
        
        if response:
            self.log_message("👋 Encerrando aplicação...")
            self.destroy()
        else:
            self.log_message("😊 Que bom que ficou! Continue usando o MDnator!")


if __name__ == "__main__":
    app = MDnatorApp()
    app.mainloop()