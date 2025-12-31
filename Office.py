import sys
import os
import subprocess
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QLabel, QComboBox, QPushButton, QTextEdit, QMessageBox, QProgressBar, QHBoxLayout, QFrame, QLineEdit, QFileDialog)
from PyQt6.QtCore import QProcess, Qt, QTimer
from PyQt6.QtGui import QIcon, QPixmap

class OfficeInstaller(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Instalador Office Pro")
        self.setGeometry(100, 100, 600, 600)
        
        # Diretório Base (Compatível com PyInstaller e Dev)
        self.basedir = self.resource_path("")
        
        if getattr(sys, 'frozen', False):
             self.assets_dir = sys._MEIPASS
             self.work_dir = os.path.dirname(sys.executable)
        else:
             self.assets_dir = os.path.dirname(os.path.abspath(__file__))
             self.work_dir = self.assets_dir

        # Carregar Ícone (do assets_dir)
        icon_path = os.path.join(self.assets_dir, "icon.png")
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))

        # Configuração do Tema Escuro
        self.apply_stylesheet()

        # Widget Central
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(0, 0, 0, 0)
        central_widget.setLayout(main_layout)

        # 1. Banner Superior
        self.header_label = QLabel()
        self.header_label.setFixedHeight(120)
        self.header_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.header_label.setStyleSheet("background-color: #2b2b2b;")
        
        banner_path = os.path.join(self.assets_dir, "banner.png")
        if os.path.exists(banner_path):
            pixmap = QPixmap(banner_path)
            self.header_label.setPixmap(pixmap.scaled(600, 120, Qt.AspectRatioMode.KeepAspectRatioByExpanding, Qt.TransformationMode.SmoothTransformation))
        else:
            self.header_label.setText("OFFICE INSTALLER")
            self.header_label.setStyleSheet("background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #ff5f6d, stop:1 #ffc371); color: white; font-size: 24px; font-weight: bold;")
        
        main_layout.addWidget(self.header_label)

        # Container de Conteúdo
        content_widget = QWidget()
        content_layout = QVBoxLayout()
        content_layout.setContentsMargins(20, 20, 20, 20)
        content_layout.setSpacing(10) # Reduzi um pouco o espaçamento
        content_widget.setLayout(content_layout)
        main_layout.addWidget(content_widget)

        # 2. Seleção de Pasta (NOVO)
        lbl_dir = QLabel("Pasta dos Arquivos (Setup e XML):")
        lbl_dir.setStyleSheet("font-size: 14px; font-weight: bold; color: #ddd;")
        content_layout.addWidget(lbl_dir)

        dir_layout = QHBoxLayout()
        self.txt_dir = QLineEdit()
        self.txt_dir.setReadOnly(True)
        self.txt_dir.setText(self.work_dir)
        self.txt_dir.setStyleSheet("padding: 5px; background: #333; color: white; border: 1px solid #555; border-radius: 4px;")
        
        self.btn_dir = QPushButton("Procurar...")
        self.btn_dir.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_dir.setFixedHeight(30)
        self.btn_dir.setStyleSheet("background: #555; color: white; border-radius: 4px; padding: 5px;")
        self.btn_dir.clicked.connect(self.select_directory)

        dir_layout.addWidget(self.txt_dir)
        dir_layout.addWidget(self.btn_dir)
        content_layout.addLayout(dir_layout)

        # 3. Seleção de Arquitetura
        lbl_arch = QLabel("Arquitetura:")
        lbl_arch.setStyleSheet("font-size: 14px; font-weight: bold; color: #ddd; margin-top: 5px;")
        content_layout.addWidget(lbl_arch)

        arch_layout = QHBoxLayout()
        self.btn_all = self.create_radio("Todos", True)
        self.btn_x64 = self.create_radio("64 Bits (x64)")
        self.btn_x86 = self.create_radio("32 Bits (x86)")
        
        arch_layout.addWidget(self.btn_all)
        arch_layout.addWidget(self.btn_x64)
        arch_layout.addWidget(self.btn_x86)
        content_layout.addLayout(arch_layout)

        # 4. Seleção de Arquivo
        lbl_select = QLabel("Selecione a Configuração:")
        lbl_select.setStyleSheet("font-size: 14px; font-weight: bold; color: #ddd; margin-top: 5px;")
        content_layout.addWidget(lbl_select)

        self.combo_xml = QComboBox()
        self.combo_xml.setStyleSheet("""
            QComboBox { padding: 8px; border-radius: 4px; border: 1px solid #555; background: #333; color: white; }
            QComboBox::drop-down { border: 0px; }
        """)
        self.populate_xml_files()
        content_layout.addWidget(self.combo_xml)

        # 5. Botão Grande
        self.btn_install = QPushButton("INICIAR INSTALAÇÃO")
        self.btn_install.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_install.setFixedHeight(50)
        self.btn_install.setStyleSheet("""
            QPushButton {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #e53935, stop:1 #e35d5b);
                color: white;
                font-weight: bold;
                font-size: 16px;
                border-radius: 5px;
                border: none;
            }
            QPushButton:hover { background: #d32f2f; }
            QPushButton:pressed { background: #b71c1c; }
            QPushButton:disabled { background: #555; color: #888; }
        """)
        self.btn_install.clicked.connect(self.run_installation)
        content_layout.addWidget(self.btn_install)

        # 6. Barra de Progresso (Indeterminado)
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setRange(0, 0) # Indeterminado
        self.progress_bar.setStyleSheet("""
            QProgressBar { border: 0px; border-radius: 4px; background-color: #444; height: 8px; }
            QProgressBar::chunk { background-color: #ff9800; border-radius: 4px; }
        """)
        content_layout.addWidget(self.progress_bar)

        # 7. Logs (Terminal Style)
        lbl_log = QLabel("Log de Atividade:")
        lbl_log.setStyleSheet("color: #aaa; margin-top: 10px;")
        content_layout.addWidget(lbl_log)


        self.log_area = QTextEdit()
        self.log_area.setReadOnly(True)
        self.log_area.setStyleSheet("""
            QTextEdit {
                background-color: #1e1e1e;
                color: #00ff00;
                font-family: Consolas, Monospace;
                border: 1px solid #333;
                border-radius: 4px;
                padding: 5px;
            }
        """)
        content_layout.addWidget(self.log_area)

        # 8. Créditos
        lbl_credits = QLabel("Desenvolvido por: Leonardo Fonseca" 
        "\nEmail:Leonardosouzafonseca902@gmail.com")
        lbl_credits.setAlignment(Qt.AlignmentFlag.AlignRight)
        lbl_credits.setStyleSheet("color: #666; font-size: 11px; margin-top: 5px; font-style: italic;")
        content_layout.addWidget(lbl_credits)

    def create_radio(self, text, checked=False):
        from PyQt6.QtWidgets import QRadioButton
        btn = QRadioButton(text)
        btn.setChecked(checked)
        btn.setStyleSheet("color: white; font-size: 12px;")
        btn.toggled.connect(self.populate_xml_files)
        return btn

    def apply_stylesheet(self):
        self.setStyleSheet("""
            QMainWindow { background-color: #202020; }
            QLabel { color: #ffffff; }
        """)

    def select_directory(self):
        new_dir = QFileDialog.getExistingDirectory(self, "Selecionar Pasta com Arquivos do Office", self.work_dir)
        if new_dir:
            self.work_dir = new_dir
            self.txt_dir.setText(self.work_dir)
            self.populate_xml_files()
            self.log(f"Diretório alterado para: {self.work_dir}")

    def populate_xml_files(self):
        # Usar diretório de trabalho (onde o .exe está), para ler os XMLs novos
        current_dir = self.work_dir 
        
        try:
            all_files = [f for f in os.listdir(current_dir) if f.endswith('.xml')]
        except FileNotFoundError:
            all_files = []

        self.combo_xml.clear()
        
        # Filtragem
        filtered_files = []
        if self.btn_x64.isChecked():
            filtered_files = [f for f in all_files if "64" in f or "x64" in f]
        elif self.btn_x86.isChecked():
            filtered_files = [f for f in all_files if "86" in f or "32" in f]
        else:
            filtered_files = all_files

        if not filtered_files:
            if not all_files:
                self.combo_xml.addItem("Nenhum arquivo XML encontrado")
            else:
                self.combo_xml.addItem("Nenhum arquivo para esta arquitetura")
            
            if hasattr(self, 'btn_install'):
                self.btn_install.setEnabled(False)
        else:
            self.combo_xml.addItems(filtered_files)
            
            if hasattr(self, 'btn_install'):
                self.btn_install.setEnabled(True)
            
            # Tentar selecionar inteligentemente
            # Prioriza 'Configuração.xml' se estiver na lista 'Todos', senão o primeiro
            default_index = self.combo_xml.findText("Configuração.xml")
            if default_index >= 0:
                self.combo_xml.setCurrentIndex(default_index)

    def log(self, message):
        self.log_area.append(f"> {message}")

    def run_installation(self):
        xml_file = self.combo_xml.currentText()
        setup_exe = "setup.exe"
        
        # Caminhos Absolutos (setup e xml estão fora do exe)
        setup_path = os.path.join(self.work_dir, setup_exe)
        conf_path = os.path.join(self.work_dir, xml_file)

        if not os.path.exists(setup_path):
            QMessageBox.critical(self, "Erro", f"O arquivo '{setup_exe}' não existe em:\n{self.work_dir}\n\nCertifique-se que o executável está na mesma pasta dos arquivos do Office.")
            return

        if not os.path.exists(conf_path):
             QMessageBox.warning(self, "Aviso", "Arquivo XML inválido.")
             return

        cmd = f'"{setup_path}" /configure "{conf_path}"'
        self.log(f"Iniciando: {cmd}")
        self.log("Aguardando processo...")
        
        self.progress_bar.setVisible(True)
        self.btn_install.setEnabled(False)

        try:
            # Usando Popen para não travar a GUI
            self.process = subprocess.Popen(cmd, shell=True)
            self.log("Processo iniciado! Verifique a janela do Office.")
            self.log(f"PID: {self.process.pid}")
            
            # Timer para verificar se o processo ainda roda (simples check)
            self.timer = QTimer()
            self.timer.timeout.connect(self.check_process)
            self.timer.start(1000)
            
        except Exception as e:
            self.log(f"Erro Crítico: {e}")
            self.progress_bar.setVisible(False)
            self.btn_install.setEnabled(True)

    def check_process(self):
        if self.process.poll() is not None:
            self.timer.stop()
            self.progress_bar.setVisible(False)
            self.btn_install.setEnabled(True)
            self.log(f"Processo finalizado com código: {self.process.returncode}")
            QMessageBox.information(self, "Concluído", "O processo de instalação foi finalizado.")

    def resource_path(self, relative_path):
        """ Retorna caminho absoluto, funcionado para dev e PyInstaller """
        try:
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")
        return os.path.join(base_path, relative_path)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = OfficeInstaller()
    window.show()
    sys.exit(app.exec())
