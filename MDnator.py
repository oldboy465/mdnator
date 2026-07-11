import sys
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
    pass  # Tratado no main

import pandas as pd
import pyreadstat

# --- IMPORTAÇÃO DO NOVO MÓDULO (MARKITDOWN MICROSOFT) ---
try:
    from markitdown import MarkItDown
    MARKITDOWN_AVAILABLE = True
except ImportError:
    MARKITDOWN_AVAILABLE = False

# --- IMPORTAÇÕES PYQT6 (NOVA INTERFACE) ---
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QTextEdit, QProgressBar, QFileDialog,
    QMessageBox, QScrollArea, QCheckBox, QFrame, QMenu, QMenuBar,
    QDialog, QGraphicsDropShadowEffect, QSizePolicy
)
from PyQt6.QtCore import Qt, pyqtSignal, QObject, QTimer
from PyQt6.QtGui import QFont, QIcon, QAction, QColor, QPalette


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
        if not documents:
            return "# ERRO: Nenhum conteúdo."
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

    def _safe_count_lines(self, file_path: str) -> int:
        encodings_to_try = ['utf-8', 'utf-8-sig', 'latin-1', 'cp1252']
        for enc in encodings_to_try:
            try:
                lines = 0
                with open(file_path, 'r', encoding=enc) as f:
                    for chunk in iter(lambda: f.read(1024 * 1024), ''):
                        lines += chunk.count('\n')
                return lines
            except UnicodeDecodeError:
                continue
            except Exception:
                raise
        raise ValueError("Falha ao decodificar arquivo com os encodings suportados.")

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

                    if selected_files is not None and full_path not in selected_files:
                        continue

                    rel_path = os.path.relpath(full_path, root_dir)
                    self.log(f"Convertendo: {rel_path}")

                    try:
                        lines = self._safe_count_lines(full_path)
                        linha_info = f" ({lines} linhas)"
                    except Exception as e:
                        linha_info = f" (Arquivo Binário/Não textual ou erro: {e})"

                    output.append(f"\n## 📄 Arquivo: `{rel_path}`{linha_info}\n\n")
                    try:
                        res = self.md.convert(full_path)
                        output.append(res.text_content)
                    except Exception as e:
                        output.append(f"> ⚠️ Erro ao converter: {e}")
                    output.append("\n\n---\n\n")

        return "".join(output)


class ProjectScanner:
    """CÉREBRO ORIGINAL DE CÓDIGO"""
    TEXT_EXTENSIONS = {'.py', '.js', '.css', '.html', '.md', '.txt', '.json', '.xml', '.sql', '.java', '.c', '.cpp', '.h', '.cs', '.sh', '.bat', '.php', '.cmd'}
    LANG_MAP = {'.py': 'python', '.js': 'javascript', '.css': 'css', '.html': 'html', '.md': 'markdown', '.json': 'json', '.sql': 'sql', '.php': 'php', '.c': 'c', '.cpp': 'cpp', '.cs': 'csharp', '.bat': 'bat', '.cmd': 'bat', '.sh': 'bash'}

    def __init__(self, log_callback=None):
        self.log = log_callback if log_callback else lambda msg: print(f"[Scanner] {msg}")
        self.ignored_dirs = {'__pycache__', '.git', '.vscode', 'node_modules', 'venv', 'env'}

    def _safe_read_text(self, file_path: str) -> Tuple[str, int]:
        encodings_to_try = ['utf-8', 'utf-8-sig', 'latin-1', 'cp1252']
        for enc in encodings_to_try:
            try:
                with open(file_path, 'r', encoding=enc) as f:
                    content = f.read()
                    line_count = content.count('\n') + (1 if content else 0)
                    return content, line_count
            except UnicodeDecodeError:
                continue
            except Exception:
                raise
        raise ValueError("Falha ao decodificar arquivo com os encodings suportados.")

    def scan_directory(self, root_dir: str, selected_files: Optional[List[str]] = None) -> str:
        self.log(f"Varrendo código (Simples): {root_dir}")
        structure = [f"# 🌳 Relatório de Código\n**Raiz:** `{root_dir}`\n\n## Estrutura\n```\n"]
        contents = ["\n\n## Conteúdo\n"]

        for dirpath, dirnames, filenames in os.walk(root_dir):
            dirnames[:] = [d for d in dirnames if d not in self.ignored_dirs]
            rel_dir = os.path.relpath(dirpath, root_dir)
            prefix = "" if rel_dir == "." else "│   " * (rel_dir.count(os.sep) + 1)

            if rel_dir != ".":
                structure.append(f"{'│   ' * rel_dir.count(os.sep)}├── 📁 {os.path.basename(dirpath)}/\n")

            for filename in filenames:
                _, ext = os.path.splitext(filename)
                if ext.lower() in self.TEXT_EXTENSIONS:
                    file_path = os.path.join(dirpath, filename)

                    if selected_files is not None and file_path not in selected_files:
                        continue

                    rel_file = os.path.relpath(file_path, root_dir)
                    structure.append(f"{prefix}├── 📄 {filename}\n")

                    try:
                        file_content, line_count = self._safe_read_text(file_path)
                        contents.append(f"\n### 📄 `{rel_file}` ({line_count} linhas)\n```{self.LANG_MAP.get(ext.lower(), '')}\n")
                        contents.append(file_content)
                    except Exception as e:
                        contents.append(f"\n### 📄 `{rel_file}` (Erro)\n```{self.LANG_MAP.get(ext.lower(), '')}\n")
                        contents.append(f"Erro de leitura ou encoding: {e}")

                    contents.append("\n```\n")

        structure.append("```\n")
        return "".join(structure) + "".join(contents)


# --- SINAIS DO PYQT PARA THREAD SAFETY ---
class AppSignals(QObject):
    log = pyqtSignal(str)
    error = pyqtSignal(str)
    save = pyqtSignal(str, str)
    progress_start = pyqtSignal()
    progress_stop = pyqtSignal()


# --- INTERFACE PYQT6 (MODERNA / LAVA THEME) ---
class MDnatorApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("MDnator v3.1 - Conversor Universal & MS Engine")
        self.resize(1100, 800)
        self.setMinimumSize(900, 750)

        # Lógica de Controle
        self.source_file_path = None
        self.source_project_path = None
        self.file_checkboxes = {}

        # Motores
        self.converter = MDConverter(log_callback=self.safe_log)
        self.project_scanner = ProjectScanner(log_callback=self.safe_log)
        self.ms_engine = MSEngine(log_callback=self.safe_log)

        # Sinais para Thread Safety
        self.signals = AppSignals()
        self.signals.log.connect(self._sync_log)
        self.signals.error.connect(self._sync_error)
        self.signals.save.connect(self._sync_save)
        self.signals.progress_start.connect(self._start_progress)
        self.signals.progress_stop.connect(self._stop_progress)

        self.apply_modern_stylesheet()
        self.init_ui()

    def apply_modern_stylesheet(self):
        # QSS com estética Lava / Magma / Glassmorphism
        # Removido o SVG que corrompia os checkboxes, substituído por cores e bordas seguras
        self.setStyleSheet("""
            QMainWindow {
                background: #0D0500; /* Fundo Obsidiana Escuro */
            }
            QWidget {
                font-family: 'Segoe UI', Arial, sans-serif;
            }
            QFrame#GlassFrame {
                background-color: rgba(255, 69, 0, 0.05); /* Transparência Lava */
                border: 1px solid rgba(255, 69, 0, 0.2);
                border-radius: 12px;
            }
            QLabel {
                color: #FFDBC2;
                background: transparent;
            }
            QLabel#TitleLabel {
                font-size: 26px;
                font-weight: 900;
                color: #FF4500; /* Laranja Red Vibrante */
            }
            QLabel#SubTitleLabel {
                font-size: 13px;
                color: #D2691E;
            }
            QLabel#SectionHeader {
                font-size: 15px;
                font-weight: bold;
                color: #FF8C00;
                padding-bottom: 5px;
            }
            QPushButton {
                background-color: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #E64A19, stop:1 #D84315);
                color: #FFFFFF;
                border: 1px solid rgba(255, 255, 255, 0.1);
                border-radius: 15px; /* Bordas mais arredondadas inspiradas nas formas ovais */
                padding: 8px 18px;
                font-weight: bold;
                font-size: 13px;
            }
            QPushButton:hover {
                background-color: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #FF5722, stop:1 #E64A19);
                border: 1px solid rgba(255, 136, 0, 0.5);
            }
            QPushButton#SecondaryBtn {
                background-color: rgba(255, 69, 0, 0.1);
                color: #FFB380;
                border: 1px solid #FF4500;
            }
            QPushButton#SecondaryBtn:hover {
                background-color: rgba(255, 69, 0, 0.25);
            }
            QTextEdit {
                background-color: rgba(0, 0, 0, 0.6);
                color: #FFB380;
                border: 1px solid rgba(255, 69, 0, 0.2);
                border-radius: 8px;
                font-family: 'Consolas', monospace;
                font-size: 13px;
                padding: 10px;
            }
            QScrollBar:vertical {
                border: none;
                background: rgba(255, 69, 0, 0.05);
                width: 10px;
                border-radius: 5px;
            }
            QScrollBar::handle:vertical {
                background: rgba(255, 69, 0, 0.3);
                min-height: 30px;
                border-radius: 5px;
            }
            QScrollBar::handle:vertical:hover {
                background: rgba(255, 69, 0, 0.5);
            }
            QMenuBar {
                background-color: #1A0A00;
                color: #FFDBC2;
                border-bottom: 1px solid rgba(255, 69, 0, 0.2);
            }
            QMenuBar::item:selected {
                background-color: rgba(255, 69, 0, 0.2);
                border-radius: 4px;
            }
            QMenu {
                background-color: #1A0A00;
                color: #FFDBC2;
                border: 1px solid #FF4500;
                border-radius: 6px;
            }
            QMenu::item:selected {
                background-color: #FF4500;
                color: white;
            }
            QMenu::item:disabled {
                color: #8B4513;
            }
            QProgressBar {
                background-color: rgba(255, 255, 255, 0.05);
                border: 1px solid rgba(255, 69, 0, 0.3);
                border-radius: 4px;
                text-align: center;
                color: transparent;
                max-height: 8px;
            }
            QProgressBar::chunk {
                background-color: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #FF8C00, stop:1 #FF0000);
                border-radius: 4px;
            }
            QScrollArea {
                background: transparent;
                border: none;
            }
            QScrollArea > QWidget > QWidget {
                background: transparent;
            }
            /* ESTILO DE CHECKBOX CORRIGIDO E FUNCIONAL */
            QCheckBox {
                color: #FFDBC2;
                font-size: 14px;
                spacing: 12px;
            }
            QCheckBox::indicator {
                width: 16px;
                height: 16px;
                border-radius: 4px;
                border: 2px solid #D2691E;
                background: rgba(0, 0, 0, 0.5);
            }
            QCheckBox::indicator:hover {
                border: 2px solid #FF8C00;
            }
            QCheckBox::indicator:checked {
                background-color: #FF4500;
                border: 2px solid #FF8C00;
            }
        """)

    def init_ui(self):
        # --- MENU BAR ---
        menubar = self.menuBar()
        
        file_menu = menubar.addMenu("Arquivo")
        self.act_open = QAction("📂 Abrir documento...", self)
        self.act_open.triggered.connect(self.on_open_file)
        file_menu.addAction(self.act_open)
        self.act_exit = QAction("❌ Sair", self)
        self.act_exit.triggered.connect(self.close)
        file_menu.addAction(self.act_exit)

        proj_menu = menubar.addMenu("Projeto")
        self.act_open_proj = QAction("🌳 Selecionar Pasta...", self)
        self.act_open_proj.triggered.connect(self.on_open_project_folder)
        proj_menu.addAction(self.act_open_proj)
        self.act_proc_proj_lc = QAction("🚀 Processar (Scanner Simples)", self)
        self.act_proc_proj_lc.triggered.connect(self.on_process_project_start)
        self.act_proc_proj_lc.setEnabled(False)
        proj_menu.addAction(self.act_proc_proj_lc)

        exec_menu = menubar.addMenu("Executar (Padrão)")
        self.act_proc_file_lc = QAction("⚡ Processar Arquivo (LangChain)", self)
        self.act_proc_file_lc.triggered.connect(self.on_process_file_start)
        self.act_proc_file_lc.setEnabled(False)
        exec_menu.addAction(self.act_proc_file_lc)

        ms_menu = menubar.addMenu("MarkItDown (MS)")
        self.act_proc_file_ms = QAction("⚡ Converter Arquivo (MS Engine)", self)
        self.act_proc_file_ms.triggered.connect(self.on_ms_file_start)
        self.act_proc_file_ms.setEnabled(False)
        ms_menu.addAction(self.act_proc_file_ms)
        self.act_proc_proj_ms = QAction("🚀 Converter Projeto (MS Engine)", self)
        self.act_proc_proj_ms.triggered.connect(self.on_ms_project_start)
        self.act_proc_proj_ms.setEnabled(False)
        ms_menu.addAction(self.act_proc_proj_ms)

        help_menu = menubar.addMenu("Ajuda")
        self.act_about = QAction("ℹ️ Sobre o MDnator...", self)
        self.act_about.triggered.connect(self.on_about)
        help_menu.addAction(self.act_about)

        # --- CENTRAL WIDGET ---
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(30, 30, 30, 30)
        main_layout.setSpacing(20)

        # --- HEADER ---
        header_layout = QHBoxLayout()
        title_lbl = QLabel("📄 MDnator v3.1")
        title_lbl.setObjectName("TitleLabel")
        sub_lbl = QLabel("Lava Edition + Microsoft MarkItDown")
        sub_lbl.setObjectName("SubTitleLabel")
        header_layout.addWidget(title_lbl)
        header_layout.addWidget(sub_lbl)
        header_layout.addStretch()
        main_layout.addLayout(header_layout)

        # --- META INFO (ORIGEM / STATUS) ---
        meta_layout = QHBoxLayout()
        meta_layout.setSpacing(20)

        src_frame = QFrame()
        src_frame.setObjectName("GlassFrame")
        self._apply_shadow(src_frame)
        src_vbox = QVBoxLayout(src_frame)
        lbl_src_title = QLabel("📥 Origem")
        lbl_src_title.setObjectName("SectionHeader")
        self.lbl_src = QLabel("Nenhum item selecionado")
        src_vbox.addWidget(lbl_src_title)
        src_vbox.addWidget(self.lbl_src)
        meta_layout.addWidget(src_frame)

        status_frame = QFrame()
        status_frame.setObjectName("GlassFrame")
        self._apply_shadow(status_frame)
        status_vbox = QVBoxLayout(status_frame)
        lbl_status_title = QLabel("📤 Status")
        lbl_status_title.setObjectName("SectionHeader")
        self.lbl_status = QLabel("Aguardando...")
        status_vbox.addWidget(lbl_status_title)
        status_vbox.addWidget(self.lbl_status)
        meta_layout.addWidget(status_frame)

        main_layout.addLayout(meta_layout)

        # --- FILES LIST (CHECKBOXES) ---
        self.files_frame = QFrame()
        self.files_frame.setObjectName("GlassFrame")
        self._apply_shadow(self.files_frame)
        files_vbox = QVBoxLayout(self.files_frame)
        
        files_header_layout = QHBoxLayout()
        lbl_files_title = QLabel("☑️ Ficheiros do Projeto (Selecione para incluir)")
        lbl_files_title.setObjectName("SectionHeader")
        files_header_layout.addWidget(lbl_files_title)
        
        btn_mark_all = QPushButton("Marcar Todos")
        btn_mark_all.setObjectName("SecondaryBtn")
        btn_mark_all.clicked.connect(self.select_all_files)
        btn_unmark_all = QPushButton("Desmarcar Todos")
        btn_unmark_all.setObjectName("SecondaryBtn")
        btn_unmark_all.clicked.connect(self.deselect_all_files)
        
        files_header_layout.addStretch()
        files_header_layout.addWidget(btn_mark_all)
        files_header_layout.addWidget(btn_unmark_all)
        files_vbox.addLayout(files_header_layout)

        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_content = QWidget()
        self.scroll_layout = QVBoxLayout(self.scroll_content)
        self.scroll_layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        self.scroll_area.setWidget(self.scroll_content)
        files_vbox.addWidget(self.scroll_area)
        
        main_layout.addWidget(self.files_frame, stretch=2)

        # --- LOG TERMINAL ---
        log_frame = QFrame()
        log_frame.setObjectName("GlassFrame")
        self._apply_shadow(log_frame)
        log_vbox = QVBoxLayout(log_frame)
        lbl_log_title = QLabel("📝 Log Terminal")
        lbl_log_title.setObjectName("SectionHeader")
        log_vbox.addWidget(lbl_log_title)
        
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        log_vbox.addWidget(self.log_text)
        
        main_layout.addWidget(log_frame, stretch=1)

        # --- PROGRESS BAR ---
        self.progress = QProgressBar()
        self.progress.setRange(0, 0) # Indeterminate
        self.progress.setVisible(False)
        main_layout.addWidget(self.progress)

    def _apply_shadow(self, widget):
        shadow = QGraphicsDropShadowEffect(self)
        shadow.setBlurRadius(25)
        # Sombra alaranjada para dar aspecto incandescente (Lava)
        shadow.setColor(QColor(255, 69, 0, 40))
        shadow.setOffset(0, 4)
        widget.setGraphicsEffect(shadow)

    # --- SINAIS E SLOTS (THREAD SAFETY) ---
    def safe_log(self, msg):
        self.signals.log.emit(msg)

    def _sync_log(self, msg):
        timestamp = datetime.now().strftime('%H:%M:%S')
        self.log_text.append(f"[{timestamp}] {msg}")
        self.log_text.moveCursor(self.log_text.textCursor().MoveOperation.End)
        self.lbl_status.setText(msg[:60] + "..." if len(msg) > 60 else msg)

    def _sync_error(self, err):
        QMessageBox.critical(self, "Erro Fatal", err)

    def _start_progress(self):
        self.progress.setVisible(True)

    def _stop_progress(self):
        self.progress.setVisible(False)

    def _sync_save(self, content, suffix):
        name = f"output{suffix}.md"
        path, _ = QFileDialog.getSaveFileName(self, "Salvar Arquivo", name, "Markdown Files (*.md);;All Files (*)")
        if path:
            try:
                with open(path, 'w', encoding='utf-8') as f:
                    f.write(content)
                self.safe_log(f"Salvo em: {path}")
                QMessageBox.information(self, "Sucesso", "Arquivo salvo com sucesso!")
            except Exception as e:
                self.safe_log(f"Erro ao salvar arquivo: {e}")
                QMessageBox.critical(self, "Erro de I/O", f"Não foi possível salvar: {e}")

    def save_file(self, content, suffix=""):
        self.signals.save.emit(content, suffix)

    # --- LÓGICA DA UI ---
    def on_open_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "Abrir Arquivo")
        if path:
            self.source_file_path = path
            self.source_project_path = None
            self.lbl_src.setText(f"ARQUIVO: {os.path.basename(path)}")
            self.update_menus(file_mode=True)
            self.clear_file_list()
            self.safe_log(f"Arquivo carregado: {path}")

    def on_open_project_folder(self):
        path = QFileDialog.getExistingDirectory(self, "Selecionar Pasta do Projeto")
        if path:
            self.source_project_path = path
            self.source_file_path = None
            self.lbl_src.setText(f"PROJETO: {os.path.basename(path)}")
            self.update_menus(file_mode=False)
            self.populate_file_list(path)
            self.safe_log(f"Pasta carregada. Processando lista de ficheiros: {path}")

    def update_menus(self, file_mode):
        self.act_proc_file_lc.setEnabled(file_mode)
        self.act_proc_proj_lc.setEnabled(not file_mode)
        self.act_proc_file_ms.setEnabled(file_mode)
        self.act_proc_proj_ms.setEnabled(not file_mode)

    def clear_file_list(self):
        while self.scroll_layout.count():
            child = self.scroll_layout.takeAt(0)
            if child.widget():
                child.widget().deleteLater()
        self.file_checkboxes.clear()

    def select_all_files(self):
        for cb in self.file_checkboxes.values():
            cb.setChecked(True)

    def deselect_all_files(self):
        for cb in self.file_checkboxes.values():
            cb.setChecked(False)

    def populate_file_list(self, root_dir):
        self.clear_file_list()
        ms_exts = {'.pptx', '.docx', '.xlsx', '.pdf', '.jpg', '.png', '.html', '.csv', '.json', '.xml', '.txt', '.md'}
        all_exts = self.project_scanner.TEXT_EXTENSIONS.union(ms_exts)
        all_files_to_render = []

        for dirpath, dirnames, filenames in os.walk(root_dir):
            dirnames[:] = [d for d in dirnames if d not in self.project_scanner.ignored_dirs]
            for filename in filenames:
                _, ext = os.path.splitext(filename)
                if ext.lower() in all_exts:
                    full_path = os.path.join(dirpath, filename)
                    rel_path = os.path.relpath(full_path, root_dir)
                    all_files_to_render.append((full_path, rel_path))

        self.safe_log(f"Encontrados {len(all_files_to_render)} ficheiros suportados. Renderizando interface...")
        self._render_checkbox_chunk(all_files_to_render, root_dir, 0, 40)

    def _render_checkbox_chunk(self, files_list, root_dir, start_idx, chunk_size):
        end_idx = min(start_idx + chunk_size, len(files_list))
        for i in range(start_idx, end_idx):
            full_path, rel_path = files_list[i]
            cb = QCheckBox(rel_path)
            cb.setChecked(True)
            self.scroll_layout.addWidget(cb)
            self.file_checkboxes[full_path] = cb

        if end_idx < len(files_list):
            QTimer.singleShot(10, lambda: self._render_checkbox_chunk(files_list, root_dir, end_idx, chunk_size))
        else:
            self.safe_log("Lista de arquivos carregada e interface atualizada com sucesso.")

    def get_selected_files(self) -> List[str]:
        return [path for path, cb in self.file_checkboxes.items() if cb.isChecked()]

    # --- PROCESSADORES (THREADS) ---
    def run_thread(self, target):
        self.signals.progress_start.emit()
        threading.Thread(target=self._wrapper, args=(target,), daemon=True).start()

    def _wrapper(self, target_func):
        try:
            target_func()
        except Exception as e:
            self.safe_log(f"ERRO FATAL: {e}")
            self.signals.error.emit(str(e))
        finally:
            self.signals.progress_stop.emit()

    def on_process_file_start(self):
        self.run_thread(self._process_file_lc)

    def _process_file_lc(self):
        docs, _ = self.converter.load_document(self.source_file_path)
        md = self.converter.format_to_markdown(docs)
        self.save_file(md)

    def on_process_project_start(self):
        self.run_thread(self._process_proj_scanner)

    def _process_proj_scanner(self):
        selected_files = self.get_selected_files()
        md = self.project_scanner.scan_directory(self.source_project_path, selected_files)
        self.save_file(md)

    def on_ms_file_start(self):
        if not MARKITDOWN_AVAILABLE:
            QMessageBox.critical(self, "Erro", "Instale a lib: pip install markitdown")
            return
        self.run_thread(self._process_file_ms)

    def _process_file_ms(self):
        md = self.ms_engine.convert_file(self.source_file_path)
        self.save_file(md, suffix="_MS")

    def on_ms_project_start(self):
        if not MARKITDOWN_AVAILABLE:
            QMessageBox.critical(self, "Erro", "Instale a lib: pip install markitdown")
            return
        self.run_thread(self._process_proj_ms)

    def _process_proj_ms(self):
        selected_files = self.get_selected_files()
        md = self.ms_engine.scan_project(self.source_project_path, selected_files)
        self.save_file(md, suffix="_ProjectMS")

    def copy_to_clipboard(self, text):
        QApplication.clipboard().setText(text)
        self.safe_log(f"📋 '{text}' copiado para a área de transferência!")

    def on_about(self):
        self.safe_log("ℹ️ Exibindo janela 'Sobre'...")
        dlg = QDialog(self)
        dlg.setWindowTitle("Sobre o MDnator v3.1")
        dlg.setFixedSize(550, 600)
        dlg.setStyleSheet(self.styleSheet())
        
        layout = QVBoxLayout(dlg)
        
        header = QFrame()
        header.setObjectName("GlassFrame")
        h_layout = QVBoxLayout(header)
        title = QLabel("📄 MDnator v3.1")
        title.setStyleSheet("font-size: 20px; font-weight: bold;")
        sub = QLabel("Lava Edition & Scanner de Projeto")
        sub.setStyleSheet("color: #D2691E;")
        h_layout.addWidget(title)
        h_layout.addWidget(sub)
        layout.addWidget(header)

        content = QFrame()
        content.setObjectName("GlassFrame")
        c_layout = QVBoxLayout(content)
        
        info = QLabel(
            "Desenvolvido por: Philipe Sampaio Lima\n"
            "Em parceria com: Claude AI (Anthropic) & Gemini\n"
            "Versão: 3.1 (Edição \"Microsoft Engine + Lava Theme\")\n"
            "Data de lançamento: Janeiro 2026\n"
            "════════════════════════════════════════\n"
            "🔧 Contatos:\n"
            "   • Email: cienciaegestao@gmail.com\n"
            "   • WhatsApp: +55 98 98250-6920\n"
            "   • GitHub: @OldBoy465\n"
            "════════════════════════════════════════\n"
            "🎯 Modo 1: Conversor de Arquivo (LangChain)\n"
            "🎯 Modo 2: Scanner de Projeto (Simples)\n"
            "🎯 Modo 3: Microsoft MarkItDown Engine (NOVO!)\n"
            "════════════════════════════════════════\n"
            "© 2026 - Todos os direitos reservados"
        )
        c_layout.addWidget(info)
        
        btn_layout = QHBoxLayout()
        btn_email = QPushButton("Copiar Email")
        btn_email.clicked.connect(lambda: self.copy_to_clipboard("cienciaegestao@gmail.com"))
        btn_zap = QPushButton("Copiar WhatsApp")
        btn_zap.clicked.connect(lambda: self.copy_to_clipboard("+55 98 98250-6920"))
        btn_layout.addWidget(btn_email)
        btn_layout.addWidget(btn_zap)
        
        c_layout.addLayout(btn_layout)
        layout.addWidget(content)

        btn_close = QPushButton("Fechar")
        btn_close.clicked.connect(dlg.accept)
        layout.addWidget(btn_close)
        
        dlg.exec()

    def closeEvent(self, event):
        self.safe_log("👋 Usuário tentou sair...")
        reply = QMessageBox.question(
            self, "Já vai embora?",
            "Tem certeza que já vai, baby? 🥺\n\nO que foi? Só me usou e jogou fora.... kkk\n\nDeseja realmente sair do MDnator?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )
        if reply == QMessageBox.StandardButton.Yes:
            self.safe_log("👋 Encerrando aplicação...")
            event.accept()
        else:
            self.safe_log("😊 Que bom que ficou! Continue usando o MDnator!")
            event.ignore()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    # Define o estilo base como Fusion para se alinhar melhor com os temas escuros do QSS
    app.setStyle('Fusion')
    window = MDnatorApp()
    window.show()
    sys.exit(app.exec())