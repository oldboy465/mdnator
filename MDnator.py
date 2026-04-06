import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
import time
import threading
from datetime import datetime
from typing import List, Tuple
import hashlib

# --- IMPORTAÇÕES DO LANGCHAIN ---
from langchain_community.document_loaders import (
    PyPDFLoader,
    Docx2txtLoader,
    TextLoader,
    BSHTMLLoader,
    CSVLoader
)
from langchain_community.document_loaders.excel import UnstructuredExcelLoader as PandasExcelLoader
from langchain_core.documents import Document
import pandas as pd
import pyreadstat


class MDConverter:
    """
    O CÉREBRO DO MDNATOR v3.0 (Modo Arquivo Único)
    Encapsula toda a lógica de carregamento e extração
    de dados usando o framework LangChain.
    """
    
    def __init__(self, log_callback=None):
        """'log_callback' é a FUNÇÃO da GUI para onde enviaremos os logs."""
        self.log = log_callback if log_callback else lambda msg: print(f"[MDConverter_LOG] {msg}")
        self.log("Cérebro 'MDConverter' (Arquivo Único) inicializado.")

    def load_document(self, file_path: str) -> Tuple[List[Document], dict]:
        """
        Método principal: Identifica, carrega e extrai texto de um arquivo.
        """
        self.log(f"Analisando: {os.path.basename(file_path)}")
        
        _, file_extension = os.path.splitext(file_path)
        file_extension = file_extension.lower()
        
        documents = []
        metadata = {}

        try:
            if file_extension == '.pdf':
                self.log("Tipo: PDF. Usando PyPDFLoader...")
                loader = PyPDFLoader(file_path)
                documents = loader.load_and_split()

            elif file_extension == '.docx':
                self.log("Tipo: DOCX. Usando Docx2txtLoader...")
                loader = Docx2txtLoader(file_path)
                documents = loader.load()

            elif file_extension in ['.txt', '.md', '.py', '.php', '.css', '.js', '.java', '.sql', '.html', '.htm', '.json', '.xml']:
                self.log(f"Tipo: Texto ({file_extension}). Usando TextLoader...")
                loader = TextLoader(file_path, encoding='utf-8')
                documents = loader.load()

            elif file_extension == '.csv':
                self.log("Tipo: CSV. Usando CSVLoader...")
                loader = CSVLoader(file_path, encoding='utf-8')
                documents = loader.load()

            elif file_extension in ['.xls', '.xlsx']:
                self.log(f"Tipo: Excel ({file_extension}). Usando PandasExcelLoader...")
                loader = PandasExcelLoader(file_path, sheet_name=None)
                documents = loader.load()

            elif file_extension == '.html' or file_extension == '.htm':
                self.log("Tipo: HTML. Usando BSHTMLLoader...")
                loader = BSHTMLLoader(file_path, open_encoding='utf-8')
                documents = loader.load()

            elif file_extension == '.sav':
                self.log("Tipo: SPSS (.sav). Usando Pandas/pyreadstat...")
                df, meta = pyreadstat.read_sav(file_path)
                md_content = df.to_markdown(index=False)
                doc = Document(
                    page_content=md_content,
                    metadata={"source": file_path, "column_labels": meta.column_labels}
                )
                documents = [doc]

            elif file_extension == '.doc':
                self.log("AVISO: .doc legado não suportado. Use .docx.")
                raise NotImplementedError("Formato .doc legado não suportado. Use .docx.")

            else:
                self.log(f"ERRO: Tipo de arquivo '{file_extension}' não suportado pelo modo de arquivo único.")
                raise ValueError(f"Tipo de arquivo não suportado: {file_extension}")

            if documents:
                metadata = documents[0].metadata
                self.log(f"Carregamento bem-sucedido. {len(documents)} 'Documento(s)' extraído(s).")
            
            return documents, metadata

        except Exception as e:
            self.log(f"ERRO CRÍTICO no carregamento: {e}")
            raise

    def format_to_markdown(self, documents: List[Document]) -> str:
        """
        Consolida a lista de Documentos do LangChain em uma
        única string de Markdown limpa.
        """
        self.log("Formatando Documentos extraídos para Markdown...")
        
        if not documents:
            return "# ERRO: Nenhum conteúdo foi extraído."

        output_md = []
        output_md.append(f"# 🧠 MDnator v3.0 - Conversão de Documento\n\n")
        source = documents[0].metadata.get('source', 'Fonte Desconhecida')
        output_md.append(f"**Arquivo de Origem:** `{os.path.basename(source)}`\n\n")
        output_md.append(f"**Data de Conversão:** {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n\n")
        output_md.append("---\n\n")

        for i, doc in enumerate(documents):
            page_content = doc.page_content.strip()
            
            if len(documents) > 1:
                page_num = doc.metadata.get('page')
                row_num = doc.metadata.get('row')
                sheet_name = doc.metadata.get('sheet_name')

                if page_num is not None:
                    output_md.append(f"## 📄 Página {page_num + 1}\n\n")
                elif row_num is not None:
                     output_md.append(f"### 📊 Linha {row_num + 1}\n\n")
                elif sheet_name is not None:
                     output_md.append(f"## 📈 Planilha: {sheet_name}\n\n")
                else:
                    output_md.append(f"## 🧩 Seção {i + 1}\n\n")
            
            output_md.append(page_content)
            output_md.append("\n\n---\n\n")
            
        self.log("Formatação para Markdown concluída.")
        return "".join(output_md)


class ProjectScanner:
    """
    O NOVO CÉREBRO DO MDNATOR v3.0 (Modo Projeto)
    Varre um diretório de projeto inteiro e o consolida em
    um único arquivo Markdown.
    """
    
    # Lista de extensões de arquivo que queremos incluir (foco em texto/código)
    # Adicione ou remova conforme necessário
    TEXT_EXTENSIONS = {
        '.py', '.js', '.css', '.html', '.htm', '.md', '.txt', '.json', '.xml',
        '.sql', '.java', '.c', '.cpp', '.h', '.cs', '.sh', '.bat', '.gitignore',
        '.dockerfile', 'readme', '.php', '.config', '.ini', '.properties'
    }
    
    # Mapeamento de extensões para linguagem do Markdown
    LANG_MAP = {
        '.py': 'python',
        '.js': 'javascript',
        '.css': 'css',
        '.html': 'html',
        '.md': 'markdown',
        '.json': 'json',
        '.xml': 'xml',
        '.sql': 'sql',
        '.java': 'java',
        '.c': 'c',
        '.cpp': 'cpp',
        '.h': 'c',
        '.cs': 'csharp',
        '.sh': 'bash',
        '.bat': 'batch',
        '.php': 'php',
    }

    def __init__(self, log_callback=None):
        self.log = log_callback if log_callback else lambda msg: print(f"[ProjectScanner_LOG] {msg}")
        self.log("Cérebro 'ProjectScanner' (Modo Projeto) inicializado.")
        self.ignored_dirs = {'__pycache__', '.git', '.vscode', 'node_modules', 'venv', 'env'}
        self.ignored_files = {'.DS_Store'}

    def _is_text_file(self, file_name: str) -> bool:
        """Verifica se a extensão do arquivo está na nossa lista de permissão."""
        if file_name in self.ignored_files:
            return False
        
        _, ext = os.path.splitext(file_name)
        return ext.lower() in self.TEXT_EXTENSIONS

    def _get_md_lang(self, file_name: str) -> str:
        """Retorna a tag de linguagem do Markdown para a extensão do arquivo."""
        _, ext = os.path.splitext(file_name)
        return self.LANG_MAP.get(ext.lower(), '')

    def scan_directory(self, root_dir: str) -> str:
        """
        Método principal: Varre o diretório e gera o relatório Markdown.
        """
        self.log(f"Iniciando varredura do projeto em: {root_dir}")
        
        file_structure = []
        file_contents = []
        
        root_name = os.path.basename(root_dir)
        
        file_structure.append(f"# 🌳 Relatório do Projeto: {root_name}\n\n")
        file_structure.append(f"**Diretório Raiz:** `{root_dir}`\n")
        file_structure.append(f"**Data da Varredura:** {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n\n")
        
        file_structure.append("## 🏗️ Estrutura de Arquivos\n\n")
        file_structure.append("```\n")
        file_structure.append(f"📁 {root_name}/\n")

        file_contents.append("\n\n---\n\n## 📜 Conteúdo dos Arquivos\n")

        total_files_scanned = 0
        
        for dirpath, dirnames, filenames in os.walk(root_dir, topdown=True):
            # --- Ignora diretórios indesejados ---
            dirnames[:] = [d for d in dirnames if d not in self.ignored_dirs]
            
            # --- Calcula o nível de profundidade para a árvore ---
            relative_dir_path = os.path.relpath(dirpath, root_dir)
            if relative_dir_path == ".":
                depth = 0
                prefix = ""
            else:
                depth = relative_dir_path.count(os.sep) + 1
                prefix = "│   " * depth
            
            # --- Adiciona diretórios à árvore ---
            if depth > 0: # Não repete o root
                dir_prefix = "│   " * (depth - 1)
                file_structure.append(f"{dir_prefix}├── 📁 {os.path.basename(dirpath)}/\n")

            # --- Processa os arquivos ---
            for i, filename in enumerate(filenames):
                if self._is_text_file(filename):
                    total_files_scanned += 1
                    file_path = os.path.join(dirpath, filename)
                    relative_file_path = os.path.relpath(file_path, root_dir)
                    
                    self.log(f"Lendo: {relative_file_path}")

                    # --- Adiciona arquivo à árvore ---
                    file_structure.append(f"{prefix}├── 📄 {filename}\n")
                    
                    # --- Adiciona conteúdo do arquivo ---
                    file_contents.append(f"\n### 📄 `{relative_file_path}`\n\n")
                    lang = self._get_md_lang(filename)
                    file_contents.append(f"```{lang}\n")
                    
                    try:
                        with open(file_path, 'r', encoding='utf-8') as f:
                            file_contents.append(f.read())
                    except Exception as e:
                        file_contents.append(f"ERRO AO LER O ARQUIVO: {e}")
                        self.log(f"❌ Erro ao ler {relative_file_path}: {e}")
                    
                    file_contents.append("\n```\n")
        
        file_structure.append("```\n")
        
        self.log(f"Varredura concluída. {total_files_scanned} arquivos de texto processados.")
        
        return "".join(file_structure) + "".join(file_contents)


class MDnatorApp(tk.Tk):
    """
    Classe principal da GUI do MDnator v3.0.
    Agora com dupla funcionalidade: Arquivo Único e Scanner de Projeto.
    """
    
    def __init__(self):
        super().__init__()
        
        # --- Configuração Inicial da Janela ---
        self.title("MDnator v3.0 - Conversor Universal & Scanner de Projeto")
        self.geometry("1000x700+100+100")
        self.minsize(900, 650)
        
        # --- Estilo Personalizado ---
        self.style = ttk.Style()
        self.style.theme_use('clam')
        
        # Cores personalizadas
        self.bg_color = "#f0f4f8"
        self.accent_color = "#2563eb"
        self.success_color = "#10b981"
        self.warning_color = "#f59e0b"
        
        self.configure(bg=self.bg_color)
        
        # Estilos customizados
        self.style.configure('Title.TLabel', font=('Segoe UI', 11, 'bold'), background=self.bg_color)
        self.style.configure('Info.TLabel', font=('Segoe UI', 9), background='white')
        self.style.configure('Header.TLabelframe.Label', font=('Segoe UI', 10, 'bold'))
        self.style.configure('Header.TLabelframe', background=self.bg_color)
        
        # --- Variáveis de Estado ---
        self.source_file_path = None
        self.source_project_path = None
        
        # --- INICIALIZAÇÃO DOS CÉREBROS ---
        self.converter = MDConverter(log_callback=self.log_message_thread_safe)
        self.project_scanner = ProjectScanner(log_callback=self.log_message_thread_safe)

        # --- Construção da GUI ---
        self.create_header()
        self.create_menu()
        self.create_widgets()
        
        # --- Gancho do Easter Egg ---
        self.protocol("WM_DELETE_WINDOW", self.on_exit_request)

    def create_header(self):
        """Cria o cabeçalho da aplicação."""
        header_frame = tk.Frame(self, bg=self.accent_color, height=70)
        header_frame.pack(fill=tk.X, side=tk.TOP)
        header_frame.pack_propagate(False)
        
        title_label = tk.Label(
            header_frame, 
            text="📄 MDnator v3.0", 
            font=('Segoe UI', 18, 'bold'),
            bg=self.accent_color,
            fg='white'
        )
        title_label.pack(side=tk.LEFT, padx=20, pady=15)
        
        subtitle_label = tk.Label(
            header_frame, 
            text="Conversor Universal e Scanner de Projeto", 
            font=('Segoe UI', 10),
            bg=self.accent_color,
            fg='white'
        )
        subtitle_label.pack(side=tk.LEFT, pady=15)

    def create_menu(self):
        """Cria a Barra de Menu superior."""
        self.menu_bar = tk.Menu(self)
        self.config(menu=self.menu_bar)
        
        # --- Menu "Arquivo" (Para arquivos únicos) ---
        file_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="Arquivo", menu=file_menu)
        file_menu.add_command(label="📂 Abrir documento...", command=self.on_open_file)
        file_menu.add_separator()
        file_menu.add_command(label="❌ Sair", command=self.on_exit_request)

        # --- NOVO Menu "Projeto" (Para pastas) ---
        self.project_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="Projeto", menu=self.project_menu)
        self.project_menu.add_command(label="🌳 Selecionar Pasta...", command=self.on_open_project_folder)
        self.project_menu.add_command(
            label="🚀 Processar Projeto", 
            command=self.on_process_project_start,
            state="disabled"
        )

        # --- Menu "Executar" (Para arquivos únicos) ---
        self.exec_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="Executar", menu=self.exec_menu)
        self.exec_menu.add_command(
            label="⚡ Processar Arquivo", 
            command=self.on_process_file_start,
            state="disabled"
        )

        # --- Menu "Informações" ---
        info_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="Informações", menu=info_menu)
        info_menu.add_command(label="ℹ️ Sobre o MDnator...", command=self.on_about)

    def create_widgets(self):
        """Cria os painéis principais da aplicação."""
        
        # Container principal
        main_container = tk.Frame(self, bg=self.bg_color)
        main_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # --- Frame Superior (Metadados) ---
        meta_frame = tk.Frame(main_container, bg=self.bg_color)
        meta_frame.pack(fill=tk.X, side=tk.TOP, pady=(0, 10))
        
        meta_frame.columnconfigure(0, weight=1)
        meta_frame.columnconfigure(1, weight=1)

        # Painel de Origem (Agora usado por Arquivo ou Projeto)
        source_meta_frame = ttk.LabelFrame(
            meta_frame, 
            text="📥 Origem (Arquivo ou Projeto)", 
            padding=15,
            style='Header.TLabelframe'
        )
        source_meta_frame.grid(row=0, column=0, sticky="nsew", padx=5)
        
        source_inner = tk.Frame(source_meta_frame, bg='white', relief=tk.FLAT, bd=1)
        source_inner.pack(fill=tk.BOTH, expand=True)
        
        self.lbl_src_nome = tk.Label(source_inner, text="📌 Nome: Nenhum item selecionado", 
                                      font=('Segoe UI', 9), bg='white', anchor='w', padx=10, pady=5)
        self.lbl_src_nome.pack(fill=tk.X)
        
        self.lbl_src_caminho = tk.Label(source_inner, text="📂 Caminho: N/A", 
                                         font=('Segoe UI', 9), bg='white', anchor='w', padx=10, pady=5)
        self.lbl_src_caminho.pack(fill=tk.X)
        
        self.lbl_src_tamanho = tk.Label(source_inner, text="💾 Tamanho: N/A", 
                                         font=('Segoe UI', 9), bg='white', anchor='w', padx=10, pady=5)
        self.lbl_src_tamanho.pack(fill=tk.X)
        
        self.lbl_src_formato = tk.Label(source_inner, text="📋 Formato: N/A", 
                                         font=('Segoe UI', 9), bg='white', anchor='w', padx=10, pady=5)
        self.lbl_src_formato.pack(fill=tk.X)
        
        self.lbl_src_criacao = tk.Label(source_inner, text="🕐 Criação: N/A", 
                                         font=('Segoe UI', 9), bg='white', anchor='w', padx=10, pady=5)
        self.lbl_src_criacao.pack(fill=tk.X)
        
        self.lbl_src_modificado = tk.Label(source_inner, text="🕑 Modificação: N/A", 
                                            font=('Segoe UI', 9), bg='white', anchor='w', padx=10, pady=5)
        self.lbl_src_modificado.pack(fill=tk.X)
        
        self.lbl_src_hash = tk.Label(source_inner, text="🔐 MD5: N/A", 
                                      font=('Segoe UI', 8), bg='white', anchor='w', padx=10, pady=5)
        self.lbl_src_hash.pack(fill=tk.X)

        # Painel de Destino
        dest_meta_frame = ttk.LabelFrame(
            meta_frame, 
            text="📤 Arquivo de Destino (Markdown)", 
            padding=15,
            style='Header.TLabelframe'
        )
        dest_meta_frame.grid(row=0, column=1, sticky="nsew", padx=5)
        
        dest_inner = tk.Frame(dest_meta_frame, bg='white', relief=tk.FLAT, bd=1)
        dest_inner.pack(fill=tk.BOTH, expand=True)

        self.lbl_dest_nome = tk.Label(dest_inner, text="📌 Nome: Aguardando conversão", 
                                       font=('Segoe UI', 9), bg='white', anchor='w', padx=10, pady=5)
        self.lbl_dest_nome.pack(fill=tk.X)
        
        self.lbl_dest_caminho = tk.Label(dest_inner, text="📂 Caminho: N/A", 
                                          font=('Segoe UI', 9), bg='white', anchor='w', padx=10, pady=5)
        self.lbl_dest_caminho.pack(fill=tk.X)
        
        self.lbl_dest_tamanho = tk.Label(dest_inner, text="💾 Tamanho: N/A", 
                                          font=('Segoe UI', 9), bg='white', anchor='w', padx=10, pady=5)
        self.lbl_dest_tamanho.pack(fill=tk.X)
        
        self.lbl_dest_criacao = tk.Label(dest_inner, text="🕐 Criação: N/A", 
                                          font=('Segoe UI', 9), bg='white', anchor='w', padx=10, pady=5)
        self.lbl_dest_criacao.pack(fill=tk.X)
        
        self.lbl_dest_modificado = tk.Label(dest_inner, text="🕑 Modificação: N/A", 
                                             font=('Segoe UI', 9), bg='white', anchor='w', padx=10, pady=5)
        self.lbl_dest_modificado.pack(fill=tk.X)
        
        self.lbl_dest_hash = tk.Label(dest_inner, text="🔐 MD5: N/A", 
                                       font=('Segoe UI', 8), bg='white', anchor='w', padx=10, pady=5)
        self.lbl_dest_hash.pack(fill=tk.X)
        
        self.lbl_dest_tempo = tk.Label(dest_inner, text="⏱️ Tempo de conversão: N/A", 
                                        font=('Segoe UI', 9), bg='white', anchor='w', padx=10, pady=5)
        self.lbl_dest_tempo.pack(fill=tk.X)
        
        # --- Frame Inferior (Logs e Progresso) ---
        bottom_frame = tk.Frame(main_container, bg=self.bg_color)
        bottom_frame.pack(fill=tk.BOTH, expand=True, side=tk.BOTTOM)

        log_frame = ttk.LabelFrame(
            bottom_frame, 
            text="📝 Log de Execução", 
            padding=10,
            style='Header.TLabelframe'
        )
        log_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        self.log_text = tk.Text(
            log_frame, 
            height=10, 
            wrap=tk.WORD, 
            state="disabled", 
            bg="#1e293b", 
            fg="#e2e8f0",
            font=('Consolas', 9),
            padx=10,
            pady=10
        )
        log_scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.config(yscrollcommand=log_scrollbar.set)
        
        log_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        # Barra de progresso estilizada
        self.progress_bar = ttk.Progressbar(
            bottom_frame, 
            orient=tk.HORIZONTAL, 
            mode='indeterminate',
            length=400
        )
        self.progress_bar.pack(fill=tk.X, pady=(5, 0))

    def calculate_file_hash(self, filepath):
        """Calcula o hash MD5 do arquivo."""
        try:
            hash_md5 = hashlib.md5()
            with open(filepath, "rb") as f:
                for chunk in iter(lambda: f.read(4096), b""):
                    hash_md5.update(chunk)
            return hash_md5.hexdigest()
        except Exception as e:
            self.log_message_thread_safe(f"Erro ao calcular hash: {e}")
            return "N/A"

    # --- Funções do Modo ARQUIVO ÚNICO ---

    def on_open_file(self):
        """Abre o seletor de arquivos e exibe os metadados."""
        file_types = [
            ("Arquivos Suportados", "*.pdf *.docx *.xls *.xlsx *.txt *.csv *.sav *.py *.html *.php *.md"),
            ("Todos os arquivos", "*.*")
        ]
        
        filepath = filedialog.askopenfilename(
            title="Selecione o arquivo de origem", 
            filetypes=file_types
        )
        
        if not filepath:
            self.log_message_thread_safe("Abertura de arquivo cancelada.")
            return

        self.source_file_path = filepath
        self.source_project_path = None # Garante que estamos no modo arquivo
        self.log_message_thread_safe(f"✅ [Modo Arquivo] Carregado: {os.path.basename(self.source_file_path)}")
        
        self.display_source_file_metadata(filepath)
        self.clear_dest_metadata()
        
        # Ativa o menu de arquivo, desativa o de projeto
        self.exec_menu.entryconfig("⚡ Processar Arquivo", state="normal")
        self.project_menu.entryconfig("🚀 Processar Projeto", state="disabled")

    def display_source_file_metadata(self, filepath):
        """Exibe os metadados do arquivo de ORIGEM."""
        try:
            stats = os.stat(filepath)
            size_mb = stats.st_size / (1024 * 1024)
            size_kb = stats.st_size / 1024
            
            # Formatação inteligente do tamanho
            if size_mb >= 1:
                size_str = f"{size_mb:.2f} MB"
            else:
                size_str = f"{size_kb:.2f} KB"
            
            mod_time = datetime.fromtimestamp(stats.st_mtime).strftime('%d/%m/%Y %H:%M:%S')
            
            # Data de criação (funciona diferente em diferentes SOs)
            try:
                create_time = datetime.fromtimestamp(stats.st_ctime).strftime('%d/%m/%Y %H:%M:%S')
            except:
                create_time = "N/A"
            
            _, ext = os.path.splitext(filepath)
            file_hash = self.calculate_file_hash(filepath)

            self.lbl_src_nome.config(text=f"📌 Nome: {os.path.basename(filepath)}")
            self.lbl_src_caminho.config(text=f"📂 Caminho: {filepath}")
            self.lbl_src_tamanho.config(text=f"💾 Tamanho: {size_str} ({stats.st_size:,} bytes)")
            self.lbl_src_formato.config(text=f"📋 Formato: {ext.upper()}")
            self.lbl_src_criacao.config(text=f"🕐 Criação: {create_time}")
            self.lbl_src_modificado.config(text=f"🕑 Modificação: {mod_time}")
            self.lbl_src_hash.config(text=f"🔐 MD5: {file_hash}")
            
        except Exception as e:
            self.log_message_thread_safe(f"❌ Erro ao ler metadados de origem: {e}")

    def on_process_file_start(self):
        """Inicia o processo de conversão de ARQUIVO em uma THREAD."""
        if not self.source_file_path:
            messagebox.showwarning("Nenhum arquivo", "Por favor, abra um arquivo primeiro.")
            return
            
        self.log_message_thread_safe(f"🚀 [Modo Arquivo] Iniciando processamento de '{os.path.basename(self.source_file_path)}'...")
        
        self.exec_menu.entryconfig("⚡ Processar Arquivo", state="disabled")
        self.progress_bar.start(10)
        
        self.clear_dest_metadata()
        self.start_time = time.time()
        
        self.worker_thread = threading.Thread(
            target=self._process_file_in_background,
            daemon=True
        )
        self.worker_thread.start()

    def _process_file_in_background(self):
        """Esta função (ARQUIVO) roda na THREAD separada."""
        try:
            documents, _ = self.converter.load_document(self.source_file_path)
            
            if not documents:
                raise Exception("Nenhum documento foi extraído. Verifique o log.")

            final_markdown = self.converter.format_to_markdown(documents)
            
            conversion_time = time.time() - self.start_time
            self.after(0, self.ask_to_save_markdown, final_markdown, conversion_time)

        except Exception as e:
            self.log_message_thread_safe(f"❌ ERRO DE PROCESSAMENTO (Arquivo): {e}")
            self.after(0, messagebox.showerror, "Erro de Conversão", 
                       f"Ocorreu um erro durante a conversão:\n\n{e}")
        
        finally:
            self.after(0, self.on_file_processing_complete)

    def on_file_processing_complete(self):
        """Limpa a barra de progresso e reativa o botão de ARQUIVO."""
        self.progress_bar.stop()
        self.exec_menu.entryconfig("⚡ Processar Arquivo", state="normal")
        self.log_message_thread_safe("✅ Processo (Arquivo) finalizado.")

    # --- Funções do NOVO Modo PROJETO ---

    def on_open_project_folder(self):
        """Abre o seletor de PASTAS."""
        folder_path = filedialog.askdirectory(
            title="Selecione a pasta raiz do projeto"
        )
        
        if not folder_path:
            self.log_message_thread_safe("Seleção de pasta cancelada.")
            return

        self.source_project_path = folder_path
        self.source_file_path = None # Garante que estamos no modo projeto
        self.log_message_thread_safe(f"✅ [Modo Projeto] Carregado: {folder_path}")
        
        self.display_source_folder_metadata(folder_path)
        self.clear_dest_metadata()
        
        # Ativa o menu de projeto, desativa o de arquivo
        self.project_menu.entryconfig("🚀 Processar Projeto", state="normal")
        self.exec_menu.entryconfig("⚡ Processar Arquivo", state="disabled")

    def display_source_folder_metadata(self, folder_path):
        """Exibe os metadados simplificados para uma PASTA."""
        self.lbl_src_nome.config(text=f"📌 Nome: {os.path.basename(folder_path)}")
        self.lbl_src_caminho.config(text=f"📂 Caminho: {folder_path}")
        self.lbl_src_tamanho.config(text="💾 Tamanho: N/A (Pasta)")
        self.lbl_src_formato.config(text="📋 Formato: DIRETÓRIO")
        self.lbl_src_criacao.config(text="🕐 Criação: N/A")
        self.lbl_src_modificado.config(text="🕑 Modificação: N/A")
        self.lbl_src_hash.config(text="🔐 MD5: N/A")

    def on_process_project_start(self):
        """Inicia o processo de varredura de PROJETO em uma THREAD."""
        if not self.source_project_path:
            messagebox.showwarning("Nenhuma pasta", "Por favor, selecione uma pasta de projeto primeiro.")
            return
            
        self.log_message_thread_safe(f"🚀 [Modo Projeto] Iniciando varredura de '{os.path.basename(self.source_project_path)}'...")
        
        self.project_menu.entryconfig("🚀 Processar Projeto", state="disabled")
        self.progress_bar.start(10)
        
        self.clear_dest_metadata()
        self.start_time = time.time()
        
        self.project_worker_thread = threading.Thread(
            target=self._process_project_in_background,
            daemon=True
        )
        self.project_worker_thread.start()

    def _process_project_in_background(self):
        """Esta função (PROJETO) roda na THREAD separada."""
        try:
            final_markdown = self.project_scanner.scan_directory(self.source_project_path)
            
            conversion_time = time.time() - self.start_time
            self.after(0, self.ask_to_save_markdown, final_markdown, conversion_time)

        except Exception as e:
            self.log_message_thread_safe(f"❌ ERRO DE VARREDURA (Projeto): {e}")
            self.after(0, messagebox.showerror, "Erro de Varredura", 
                       f"Ocorreu um erro durante a varredura do projeto:\n\n{e}")
        
        finally:
            self.after(0, self.on_project_processing_complete)

    def on_project_processing_complete(self):
        """Limpa a barra de progresso e reativa o botão de PROJETO."""
        self.progress_bar.stop()
        self.project_menu.entryconfig("🚀 Processar Projeto", state="normal")
        self.log_message_thread_safe("✅ Processo (Projeto) finalizado.")


    # --- Funções COMPARTILHADAS e da GUI ---

    def clear_dest_metadata(self):
        """Limpa os metadados do arquivo de DESTINO."""
        self.lbl_dest_nome.config(text="📌 Nome: Aguardando processamento")
        self.lbl_dest_caminho.config(text="📂 Caminho: N/A")
        self.lbl_dest_tamanho.config(text="💾 Tamanho: N/A")
        self.lbl_dest_criacao.config(text="🕐 Criação: N/A")
        self.lbl_dest_modificado.config(text="🕑 Modificação: N/A")
        self.lbl_dest_hash.config(text="🔐 MD5: N/A")
        self.lbl_dest_tempo.config(text="⏱️ Tempo de processamento: N/A")

    def display_dest_metadata(self, filepath, conversion_time):
        """Exibe os metadados do arquivo de DESTINO (o .md salvo)."""
        try:
            stats = os.stat(filepath)
            size_kb = stats.st_size / 1024
            size_bytes = stats.st_size
            
            mod_time = datetime.fromtimestamp(stats.st_mtime).strftime('%d/%m/%Y %H:%M:%S')
            
            try:
                create_time = datetime.fromtimestamp(stats.st_ctime).strftime('%d/%m/%Y %H:%M:%S')
            except:
                create_time = "N/A"
            
            file_hash = self.calculate_file_hash(filepath)

            self.lbl_dest_nome.config(text=f"📌 Nome: {os.path.basename(filepath)}")
            self.lbl_dest_caminho.config(text=f"📂 Caminho: {filepath}")
            self.lbl_dest_tamanho.config(text=f"💾 Tamanho: {size_kb:.2f} KB ({size_bytes:,} bytes)")
            self.lbl_dest_criacao.config(text=f"🕐 Criação: {create_time}")
            self.lbl_dest_modificado.config(text=f"🕑 Modificação: {mod_time}")
            self.lbl_dest_hash.config(text=f"🔐 MD5: {file_hash}")
            self.lbl_dest_tempo.config(text=f"⏱️ Tempo de processamento: {conversion_time:.2f} segundos")
            
        except Exception as e:
            self.log_message_thread_safe(f"❌ Erro ao ler metadados de destino: {e}")
            
    def ask_to_save_markdown(self, md_content: str, conversion_time: float):
        """Abre a caixa de diálogo 'Salvar como...' (Usada por ambos os modos)."""
        
        # Define um nome padrão baseado no modo
        if self.source_project_path:
            base_name = os.path.basename(self.source_project_path)
            default_name = f"Projeto_{base_name}.md"
        elif self.source_file_path:
            base_name = os.path.basename(self.source_file_path)
            default_name = os.path.splitext(base_name)[0] + ".md"
        else:
            default_name = "conversao.md"
        
        save_path = filedialog.asksaveasfilename(
            title="Salvar arquivo Markdown como...",
            initialfile=default_name,
            filetypes=[("Markdown", "*.md"), ("Todos os arquivos", "*.*")],
            defaultextension=".md"
        )
        
        if save_path:
            try:
                with open(save_path, 'w', encoding='utf-8') as f:
                    f.write(md_content)
                self.log_message_thread_safe(f"✅ SUCESSO! Arquivo salvo em: {save_path}")
                self.display_dest_metadata(save_path, conversion_time)
                messagebox.showinfo(
                    "Processamento Concluído", 
                    f"Arquivo gerado com sucesso!\n\n"
                    f"Tempo: {conversion_time:.2f}s\n"
                    f"Local: {save_path}"
                )
            except Exception as e:
                self.log_message_thread_safe(f"❌ Erro ao salvar arquivo: {e}")
                messagebox.showerror("Erro ao Salvar", f"Não foi possível salvar o arquivo:\n{e}")
        else:
            self.log_message_thread_safe("⚠️ Salvamento cancelado pelo usuário.")

    def log_message_thread_safe(self, message: str):
        """Adiciona mensagem ao log de forma thread-safe."""
        self.after(0, self.update_log_text_widget, message)

    def update_log_text_widget(self, message: str):
        """Atualiza o widget de log."""
        self.log_text.config(state="normal")
        timestamp = datetime.now().strftime('%H:%M:%S')
        
        # Colorização por tipo de mensagem
        if "✅" in message or "SUCESSO" in message:
            tag = "success"
        elif "❌" in message or "ERRO" in message:
            tag = "error"
        elif "⚠️" in message or "AVISO" in message:
            tag = "warning"
        else:
            tag = "info"
        
        self.log_text.tag_config("success", foreground="#10b981")
        self.log_text.tag_config("error", foreground="#ef4444")
        self.log_text.tag_config("warning", foreground="#f59e0b")
        self.log_text.tag_config("info", foreground="#e2e8f0")
        
        self.log_text.insert(tk.END, f"[{timestamp}] ", "info")
        self.log_text.insert(tk.END, f"{message}\n", tag)
        self.log_text.see(tk.END)
        self.log_text.config(state="disabled")

    def copy_to_clipboard(self, text):
        """Copia texto para a área de transferência."""
        self.clipboard_clear()
        self.clipboard_append(text)
        self.update() 
        self.log_message_thread_safe(f"📋 '{text}' copiado para a área de transferência!")

    def create_contact_frame(self, parent, label, text, button_text, command):
        """Cria um frame de contato com botão de copiar."""
        frame = tk.Frame(parent, bg='white')
        frame.pack(fill=tk.X, pady=2)
        
        label_widget = tk.Label(
            frame, 
            text=label, 
            font=('Segoe UI', 9), 
            bg='white', 
            anchor='w'
        )
        label_widget.pack(side=tk.LEFT, padx=(10, 5))
        
        text_widget = tk.Label(
            frame, 
            text=text, 
            font=('Segoe UI', 9), 
            bg='white', 
            fg='#334155',
            anchor='w'
        )
        text_widget.pack(side=tk.LEFT, padx=(0, 10))
        
        copy_btn = tk.Button(
            frame,
            text=button_text,
            command=command,
            bg='#e2e8f0',
            fg='#334155',
            font=('Segoe UI', 8),
            relief=tk.FLAT,
            padx=8,
            pady=2,
            cursor='hand2'
        )
        copy_btn.pack(side=tk.RIGHT, padx=(0, 10))

    def create_animated_exit(self):
        """Cria animação de tela piscando antes de fechar."""
        # Cria uma janela fullscreen preta
        self.animation_window = tk.Toplevel(self)
        self.animation_window.attributes('-fullscreen', True)
        self.animation_window.configure(bg='black')
        self.animation_window.attributes('-alpha', 0.0)  # Começa transparente
        
        # Força a janela ficar no topo
        self.animation_window.attributes('-topmost', True)
        
        def animate_flash():
            """Animação de piscar rapidamente."""
            for i in range(6):  # 6 transições (3 piscadas completas)
                alpha = 1.0 if i % 2 == 0 else 0.0
                self.animation_window.attributes('-alpha', alpha)
                self.animation_window.update()
                time.sleep(0.15)  # Intervalo entre piscadas
            
            # Fecha tudo
            self.animation_window.destroy()
            self.destroy()
        
        # Inicia a animação em uma thread separada
        threading.Thread(target=animate_flash, daemon=True).start()

    def on_about(self):
        """Exibe a janela 'Sobre' com os créditos."""
        self.log_message_thread_safe("ℹ️ Exibindo janela 'Sobre'...")
        
        about_window = tk.Toplevel(self)
        about_window.title("Sobre o MDnator v3.0")
        about_window.geometry("550x650")
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
            text="📄 MDnator v3.0",
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
Em parceria com: Gemini AI

Versão: 3.0 (Edição "Project Scan")
Data de lançamento: Novembro 2025

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

📧 Contatos:
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

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

🎯 Modo 1: Conversor de Arquivo
   • PDF, DOCX, TXT, MD
   • XLS, XLSX, CSV
   • HTML, HTM
   • SPSS (.sav)

🎯 Modo 2: Scanner de Projeto
   • Varre pastas e subpastas
   • Extrai .py, .js, .css, .sql, .html, etc.
   • Gera um .md único com a árvore e o código.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

© 2025 - Todos os direitos reservados
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
        """O EASTER EGG (Versão personalizada)."""
        self.log_message_thread_safe("👋 Usuário tentou sair...")
        
        response = messagebox.askyesno(
            "Já vai embora?", 
            "Tem certeza que já vai, baby? 🥺\n\n"
            "O que foi? Só me usou e jogou fora.... kkk\n\n"
            "Deseja realmente sair do MDnator?",
            icon='question'
        )
        
        if response:
            self.log_message_thread_safe("👋 Iniciando animação de saída...")
            self.log_message_thread_safe("💥 Tela piscando! Simulando bug...")
            
            # Desabilita a janela principal durante a animação
            self.attributes('-disabled', True)
            
            # Inicia a animação de saída
            self.create_animated_exit()
        else:
            self.log_message_thread_safe("😊 Que bom que ficou! Continue usando o MDnator!")


#
# ==============================================================================
# PONTO DE ENTRADA DA APLICAÇÃO
# ==============================================================================
#
if __name__ == "__main__":
    try:
        app = MDnatorApp()
        app.mainloop()
    except Exception as e:
        print(f"ERRO CRÍTICO AO INICIAR A APLICAÇÃO: {e}")
        try:
            root_err = tk.Tk()
            root_err.withdraw()
            messagebox.showerror(
                "Erro Fatal do MDnator", 
                f"Ocorreu um erro crítico ao iniciar:\n\n{e}\n\n"
                f"Verifique se todas as dependências estão instaladas:\n"
                f"• langchain-community\n"
                f"• pandas\n"
                f"• pyreadstat\n"
                f"• beautifulsoup4\n"
                f"• pypdf\n"
                f"• docx2txt\n"
                f"• openpyxl"
            )
        except:
            pass
