import sys
import os
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
    QFileDialog, QLabel, QMessageBox, QFrame, QSizePolicy, QSpacerItem,
    QProgressBar
)
from PyQt5.QtGui import QFont, QColor, QPalette
from PyQt5.QtCore import Qt, QSize, QPoint
import qtawesome as qta

from html_reader import transformar_planilha
from compare_movements import cruzar_planilhas_movimentacao

# Paleta de cores moderna - Tema Escuro
CORES = {
    'fundo': "#1A1B1E",
    'fundo_secundario': "#2A2B2E",
    'primaria': "#00E5FF",
    'primaria_hover': "#33EEFF",
    'texto': "#FFFFFF",
    'texto_secundario': "#A0A0A0",
    'borda': "#3A3B3E",
    'sucesso': "#00E676",
    'erro': "#FF1744",
    'aviso': "#FFD600"
}

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("CaixaSync - Processador de Planilhas BPO")
        self.setWindowFlags(Qt.FramelessWindowHint)  # Remove a barra de título padrão
        self.setAttribute(Qt.WA_TranslucentBackground)  # Permite transparência
        self.oldPos = None  # Para controlar o arrasto da janela
        
        self.setStyleSheet(f"""
            QWidget {{
                background-color: {CORES['fundo']};
                color: {CORES['texto']};
                font-family: 'Segoe UI', Arial, sans-serif;
            }}
            QLabel#titleLabel {{
                font-size: 28px;
                font-weight: bold;
                color: {CORES['primaria']};
                padding: 20px 0;
            }}
            QLabel#subtitleLabel {{
                font-size: 16px;
                color: {CORES['texto_secundario']};
                margin-bottom: 30px;
            }}
            QLabel#windowTitle {{
                font-size: 14px;
                font-weight: bold;
                color: {CORES['texto']};
            }}
            QPushButton#windowButton {{
                background-color: transparent;
                color: white;
                padding: 4px;
                min-width: 25px;
                max-width: 25px;
                min-height: 25px;
                max-height: 25px;
                border-radius: 0px;
                font-size: 14px;
                text-align: center;
                qproperty-alignment: AlignCenter;
            }}
            QPushButton#closeButton {{
                background-color: transparent;
                color: white;
                padding: 2px;
                min-width: 20px;
                max-width: 20px;
                min-height: 20px;
                max-height: 20px;
                border-radius: 0px;
                font-size: 12px;
                text-align: center;
                qproperty-alignment: AlignCenter;
            }}
            QPushButton#windowButton:hover {{
                background-color: {CORES['fundo_secundario']};
            }}
            QPushButton#closeButton:hover {{
                background-color: {CORES['erro']};
                color: {CORES['texto']};
            }}
            QLabel#statusLabel {{
                font-size: 14px;
                color: {CORES['texto_secundario']};
                min-height: 30px;
                padding: 8px 16px;
                border-radius: 6px;
                background-color: {CORES['fundo_secundario']};
            }}
            QLabel#statusLabel[status="success"] {{
                background-color: {CORES['sucesso'] + '20'};
                color: {CORES['sucesso']};
            }}
            QLabel#statusLabel[status="error"] {{
                background-color: {CORES['erro'] + '20'};
                color: {CORES['erro']};
            }}
            QPushButton {{
                background-color: {CORES['primaria']};
                color: {CORES['fundo']};
                font-size: 14px;
                font-weight: 600;
                border-radius: 8px;
                padding: 12px 24px;
                min-width: 200px;
                border: none;
            }}
            QPushButton::icon {{
                padding-right: 8px;
            }}
            QPushButton:hover {{
                background-color: {CORES['primaria_hover']};
            }}
            QPushButton:disabled {{
                background-color: {CORES['texto_secundario']};
                opacity: 0.5;
            }}
            QPushButton#secondaryButton {{
                background-color: transparent;
                color: {CORES['primaria']};
                border: 2px solid {CORES['primaria']};
            }}
            QPushButton#secondaryButton:hover {{
                background-color: {CORES['primaria'] + '15'};
            }}
            QFrame#line {{
                background-color: {CORES['borda']};
                min-height: 2px;
                max-height: 2px;
                border-radius: 1px;
            }}
            QFrame#container {{
                background-color: {CORES['fundo_secundario']};
                border-radius: 12px;
                padding: 24px;
                margin: 8px;
            }}
            QProgressBar {{
                border: none;
                border-radius: 6px;
                text-align: center;
                height: 12px;
                font-size: 12px;
                background-color: {CORES['fundo']};
            }}
            QProgressBar::chunk {{
                background-color: {CORES['primaria']};
                border-radius: 6px;
            }}
            QLabel#fileLabel {{
                background-color: {CORES['fundo']};
                padding: 8px 16px;
                border-radius: 6px;
                color: {CORES['texto_secundario']};
            }}
        """)
        self.setMinimumWidth(900)
        self.setMinimumHeight(700)
        
        # Configurar ícones com a nova cor primária
        self.icons = {
            'file': qta.icon('fa5s.file-alt', color='black'),
            'folder': qta.icon('fa5s.folder-open', color='black'),
            'process': qta.icon('fa5s.play', color=CORES['fundo']),
            'back': qta.icon('fa5s.arrow-left', color=CORES['primaria'])
        }
        
        self.initUI()
        
    def initUI(self):
        # Layout principal
        self.main_layout = QVBoxLayout()
        self.main_layout.setContentsMargins(0, 0, 0, 0)
        self.setLayout(self.main_layout)
        
        # Barra de título personalizada
        title_bar = QFrame()
        title_bar.setObjectName("titleBar")
        title_bar.setStyleSheet(f"""
            QFrame#titleBar {{
                background-color: {CORES['fundo']};
                border-bottom: 2px solid {CORES['borda']};
            }}
        """)
        title_bar.setFixedHeight(40)
        title_layout = QHBoxLayout(title_bar)
        title_layout.setContentsMargins(10, 0, 0, 0)
        title_layout.setSpacing(0)  # Remove o espaçamento entre os botões
        
        # Ícone e título
        title_label = QLabel("CaixaSync - Processador de Planilhas BPO")
        title_label.setObjectName("windowTitle")
        title_layout.addWidget(title_label)
        
        # Botões da janela
        title_layout.addStretch()
        
        minimize_btn = QPushButton("─")
        minimize_btn.setObjectName("windowButton")
        minimize_btn.clicked.connect(self.showMinimized)
        minimize_btn.setStyleSheet(f"color: black;")
        
        close_btn = QPushButton("✕")
        close_btn.setObjectName("closeButton")
        close_btn.clicked.connect(self.close)
        
        title_layout.addWidget(minimize_btn)
        title_layout.addWidget(close_btn)
        
        self.main_layout.addWidget(title_bar)
        
        # Container para o conteúdo
        content_container = QWidget()
        content_layout = QVBoxLayout(content_container)
        content_layout.setContentsMargins(20, 20, 20, 20)
        self.main_layout.addWidget(content_container)
        
        # Layout do conteúdo (antigo main_layout)
        self.content_layout = content_layout
        self.show_etapa1()

    def limpar_layout(self):
        while self.content_layout.count():
            item = self.content_layout.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()
            elif item.layout():
                self._clean_layout(item.layout())
    
    def _clean_layout(self, layout):
        while layout.count():
            item = layout.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()
            elif item.layout():
                self._clean_layout(item.layout())

    def add_separator(self):
        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setFrameShadow(QFrame.Sunken)
        line.setObjectName("line")
        self.content_layout.addWidget(line)

    def add_spacer(self, height=12):
        self.content_layout.addSpacerItem(QSpacerItem(0, height))

    def show_etapa1(self):
        self.limpar_layout()
        self.etapa1_infile = ''
        self.etapa1_outfolder = ''
        
        # Título e subtítulo
        label_titulo = QLabel("Etapa 1: Transformação da Planilha")
        label_titulo.setObjectName("titleLabel")
        label_titulo.setAlignment(Qt.AlignLeft)
        self.content_layout.addWidget(label_titulo)
        
        label_subtitulo = QLabel("Selecione a planilha HTML desformatada e a pasta onde deseja salvar o resultado")
        label_subtitulo.setObjectName("subtitleLabel")
        label_subtitulo.setAlignment(Qt.AlignLeft)
        self.content_layout.addWidget(label_subtitulo)
        
        self.add_separator()

        # Container para os inputs
        container = QFrame()
        container.setObjectName("container")
        container_layout = QVBoxLayout()
        container_layout.setSpacing(24)
        container.setLayout(container_layout)

        # INPUT arquivo
        hbox1 = QHBoxLayout()
        btn_infile = QPushButton("Selecionar Planilha")
        btn_infile.setIcon(self.icons['file'])
        btn_infile.setIconSize(QSize(20, 20))
        btn_infile.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        btn_infile.clicked.connect(self.select_etapa1_infile)
        self.infile_label = QLabel("Nenhum arquivo selecionado")
        self.infile_label.setObjectName("fileLabel")
        self.infile_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        hbox1.addWidget(btn_infile)
        hbox1.addSpacing(16)
        hbox1.addWidget(self.infile_label)
        container_layout.addLayout(hbox1)
        
        # INPUT pasta saída
        hbox2 = QHBoxLayout()
        btn_outfolder = QPushButton("Pasta de Saída")
        btn_outfolder.setIcon(self.icons['folder'])
        btn_outfolder.setIconSize(QSize(20, 20))
        btn_outfolder.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        btn_outfolder.clicked.connect(self.select_etapa1_outfolder)
        self.outfolder_label = QLabel("Nenhuma pasta selecionada")
        self.outfolder_label.setObjectName("fileLabel")
        self.outfolder_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        hbox2.addWidget(btn_outfolder)
        hbox2.addSpacing(16)
        hbox2.addWidget(self.outfolder_label)
        container_layout.addLayout(hbox2)

        # Barra de progresso
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        container_layout.addWidget(self.progress_bar)

        # Mensagem status
        self.status_label = QLabel('')
        self.status_label.setObjectName("statusLabel")
        self.status_label.setAlignment(Qt.AlignLeft)
        container_layout.addWidget(self.status_label)

        # Botão Rodar
        hbox_run = QHBoxLayout()
        btn_run1 = QPushButton("Iniciar Transformação")
        btn_run1.setIcon(self.icons['process'])
        btn_run1.setIconSize(QSize(20, 20))
        btn_run1.clicked.connect(self.run_etapa1)
        btn_run1.setLayoutDirection(Qt.LeftToRight)  # Garante direção correta do layout
        hbox_run.addStretch(1)
        hbox_run.addWidget(btn_run1, alignment=Qt.AlignCenter)  # Centraliza o botão
        hbox_run.addStretch(1)
        container_layout.addLayout(hbox_run)
        
        self.content_layout.addWidget(container)
        self.content_layout.addStretch(1)

    def show_etapa2(self):
        self.limpar_layout()
        self.etapa2_formatada = ''
        self.etapa2_movfile = ''
        self.etapa2_outfolder = ''
        
        # Título e subtítulo
        label_titulo = QLabel("Etapa 2: Cruzamento de Movimentações")
        label_titulo.setObjectName("titleLabel")
        label_titulo.setAlignment(Qt.AlignLeft)
        self.content_layout.addWidget(label_titulo)
        
        label_subtitulo = QLabel("Selecione as planilhas para cruzar e a pasta onde deseja salvar os resultados")
        label_subtitulo.setObjectName("subtitleLabel")
        label_subtitulo.setAlignment(Qt.AlignLeft)
        self.content_layout.addWidget(label_subtitulo)
        
        self.add_separator()

        # Container para os inputs
        container = QFrame()
        container.setObjectName("container")
        container_layout = QVBoxLayout()
        container_layout.setSpacing(24)
        container.setLayout(container_layout)

        # INPUT planilha formatada
        hbox0 = QHBoxLayout()
        btn_formatada = QPushButton("Planilha Formatada")
        btn_formatada.setIcon(self.icons['file'])
        btn_formatada.setIconSize(QSize(20, 20))
        btn_formatada.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        btn_formatada.clicked.connect(self.select_etapa2_formatada)
        self.formatada_label = QLabel("Nenhum arquivo selecionado")
        self.formatada_label.setObjectName("fileLabel")
        self.formatada_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        hbox0.addWidget(btn_formatada)
        hbox0.addSpacing(16)
        hbox0.addWidget(self.formatada_label)
        container_layout.addLayout(hbox0)

        # INPUT planilha movimentações
        hbox1 = QHBoxLayout()
        btn_movfile = QPushButton("Planilha de Movimentações")
        btn_movfile.setIcon(self.icons['file'])
        btn_movfile.setIconSize(QSize(20, 20))
        btn_movfile.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        btn_movfile.clicked.connect(self.select_etapa2_movfile)
        self.movfile_label = QLabel("Nenhum arquivo selecionado")
        self.movfile_label.setObjectName("fileLabel")
        self.movfile_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        hbox1.addWidget(btn_movfile)
        hbox1.addSpacing(16)
        hbox1.addWidget(self.movfile_label)
        container_layout.addLayout(hbox1)

        # INPUT pasta saída
        hbox2 = QHBoxLayout()
        btn_outfolder = QPushButton("Pasta de Saída")
        btn_outfolder.setIcon(self.icons['folder'])
        btn_outfolder.setIconSize(QSize(20, 20))
        btn_outfolder.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        btn_outfolder.clicked.connect(self.select_etapa2_outfolder)
        self.etapa2_outfolder_label = QLabel("Nenhuma pasta selecionada")
        self.etapa2_outfolder_label.setObjectName("fileLabel")
        self.etapa2_outfolder_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        hbox2.addWidget(btn_outfolder)
        hbox2.addSpacing(16)
        hbox2.addWidget(self.etapa2_outfolder_label)
        container_layout.addLayout(hbox2)

        # Barra de progresso
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        container_layout.addWidget(self.progress_bar)

        # Mensagem status
        self.status_label = QLabel('')
        self.status_label.setObjectName("statusLabel")
        self.status_label.setAlignment(Qt.AlignLeft)
        container_layout.addWidget(self.status_label)

        # Botões
        hbox_buttons = QHBoxLayout()
        
        # Botão Voltar
        btn_voltar = QPushButton("Voltar")
        btn_voltar.setObjectName("secondaryButton")
        btn_voltar.setIcon(self.icons['back'])
        btn_voltar.setIconSize(QSize(20, 20))
        btn_voltar.clicked.connect(self.show_etapa1)
        
        # Botão Rodar
        btn_run2 = QPushButton("Iniciar Cruzamento")
        btn_run2.setIcon(self.icons['process'])
        btn_run2.setIconSize(QSize(20, 20))
        btn_run2.clicked.connect(self.run_etapa2)
        
        hbox_buttons.addWidget(btn_voltar)
        hbox_buttons.addStretch(1)
        hbox_buttons.addWidget(btn_run2)
        container_layout.addLayout(hbox_buttons)
        
        self.content_layout.addWidget(container)
        self.content_layout.addStretch(1)

    def select_etapa1_infile(self):
        file, _ = QFileDialog.getOpenFileName(
            self, "Selecione a Planilha HTML desformatada",
            "", "Planilhas (*.xlsx *.xls *.csv)")
        if file:
            self.etapa1_infile = file
            self.infile_label.setText(os.path.basename(file))
            self.status_label.setText("")
        else:
            self.etapa1_infile = ''
            self.status_label.setText("Nenhum arquivo foi selecionado.")

    def select_etapa1_outfolder(self):
        folder = QFileDialog.getExistingDirectory(self, "Selecione a pasta de saída")
        if folder:
            self.etapa1_outfolder = folder
            self.outfolder_label.setText(folder)
            self.status_label.setText("")
        else:
            self.etapa1_outfolder = ''
            self.status_label.setText("Nenhuma pasta foi selecionada.")

    def run_etapa1(self):
        if not self.etapa1_infile or not self.etapa1_outfolder:
            self.status_label.setText("Selecione o arquivo e a pasta de saída.")
            QMessageBox.critical(self, "Erro", "Selecione o arquivo e a pasta de saída.")
            return
        try:
            nome_arquivo_saida = 'Planilha Formatada.xlsx'
            caminho_saida = os.path.join(self.etapa1_outfolder, nome_arquivo_saida)
            transformar_planilha(self.etapa1_infile, caminho_saida)
            self.status_label.setText("✅ Planilha transformada com sucesso!")
            QMessageBox.information(self, "Sucesso", "Planilha transformada com sucesso!")
            self.show_etapa2()
        except Exception as e:
            self.status_label.setText(f"❌ Erro ao transformar: {e}")
            QMessageBox.critical(self, "Erro", f"Erro ao transformar: {e}")

    def select_etapa2_formatada(self):
        file, _ = QFileDialog.getOpenFileName(
            self, "Selecione a Planilha Formatada",
            "", "Planilhas (*.xlsx *.xls *.csv)")
        if file:
            self.etapa2_formatada = file
            self.formatada_label.setText(os.path.basename(file))
            self.status_label.setText("")
        else:
            self.etapa2_formatada = ''
            self.status_label.setText("Nenhum arquivo foi selecionado.")

    def select_etapa2_movfile(self):
        file, _ = QFileDialog.getOpenFileName(
            self, "Selecione a Planilha de Movimentações",
            "", "Planilhas (*.xlsx *.xls *.csv)")
        if file:
            self.etapa2_movfile = file
            self.movfile_label.setText(os.path.basename(file))
            self.status_label.setText("")
        else:
            self.etapa2_movfile = ''
            self.status_label.setText("Nenhum arquivo foi selecionado.")

    def select_etapa2_outfolder(self):
        folder = QFileDialog.getExistingDirectory(self, "Pasta para salvar comparações")
        if folder:
            self.etapa2_outfolder = folder
            self.etapa2_outfolder_label.setText(folder)
            self.status_label.setText("")
        else:
            self.etapa2_outfolder = ''
            self.status_label.setText("Nenhuma pasta foi selecionada.")

    def run_etapa2(self):
        if not self.etapa2_movfile or not self.etapa2_outfolder:
            self.status_label.setText("Selecione o arquivo de movimentações e a pasta de saída.")
            QMessageBox.critical(self, "Erro", "Selecione o arquivo de movimentações e a pasta de saída.")
            return

        try:
            cruzar_planilhas_movimentacao(
                self.etapa2_formatada,
                self.etapa2_movfile,
                self.etapa2_outfolder
            )
            self.status_label.setText("✅ Comparação concluída com sucesso!")
            QMessageBox.information(self, "Sucesso", "Comparação concluída com sucesso!")
        except Exception as e:
            self.status_label.setText(f"❌ Erro ao comparar: {e}")
            QMessageBox.critical(self, "Erro", f"Erro ao comparar: {e}")

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.oldPos = event.globalPos()

    def mouseMoveEvent(self, event):
        if self.oldPos is not None:
            delta = QPoint(event.globalPos() - self.oldPos)
            self.move(self.pos() + delta)
            self.oldPos = event.globalPos()

    def mouseReleaseEvent(self, event):
        self.oldPos = None 