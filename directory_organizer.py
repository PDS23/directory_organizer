import os
import shutil
import pandas as pd
from datetime import datetime
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QLabel, QPushButton, QFileDialog,
    QTextEdit, QLineEdit, QDialog, QCheckBox, QListWidget, QListWidgetItem,
    QHBoxLayout, QMessageBox
)
from PyQt5.QtCore import Qt

class CategoriaSelector(QDialog):
    def __init__(self, categorias, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Selecione as Categorias")
        self.categorias = sorted(categorias)
        self.selecionadas = []

        layout = QVBoxLayout(self)

        self.lista = QListWidget()
        for cat in self.categorias:
            item = QListWidgetItem(cat)
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            item.setCheckState(Qt.Unchecked)
            self.lista.addItem(item)

        btn_layout = QHBoxLayout()
        btn_ok = QPushButton("OK")
        btn_cancel = QPushButton("Cancelar")

        btn_ok.clicked.connect(self.ok)
        btn_cancel.clicked.connect(self.cancel)

        btn_layout.addWidget(btn_ok)
        btn_layout.addWidget(btn_cancel)

        layout.addWidget(self.lista)
        layout.addLayout(btn_layout)

    def ok(self):
        self.selecionadas = [
            self.lista.item(i).text()
            for i in range(self.lista.count())
            if self.lista.item(i).checkState() == Qt.Checked
        ]
        if not self.selecionadas:
            QMessageBox.warning(self, "Erro", "Selecione ao menos uma categoria.")
            return
        self.accept()

    def cancel(self):
        self.reject()

class OrganizadorArquivos(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Organizador de Arquivos por Unidade > OP > OS")
        self.setGeometry(100, 100, 700, 500)

        self.planilha_path = ""
        self.origem_dir = ""
        self.destino_dir = ""
        self.categorias_selecionadas = []

        self.layout = QVBoxLayout()

        self.label_planilha = QLabel("Planilha: [não selecionada]")
        self.btn_planilha = QPushButton("Selecionar Planilha")
        self.btn_planilha.clicked.connect(self.selecionar_planilha)

        self.label_origem = QLabel("Pasta de origem: [não selecionada]")
        self.btn_origem = QPushButton("Selecionar Pasta com Arquivos")
        self.btn_origem.clicked.connect(self.selecionar_origem)

        self.label_destino = QLabel("Destino: [não selecionado]")
        self.btn_destino = QPushButton("Selecionar Pasta de Destino")
        self.btn_destino.clicked.connect(self.selecionar_destino)

        self.btn_processar = QPushButton("Processar Arquivos")
        self.btn_processar.clicked.connect(self.processar)
        self.btn_processar.setStyleSheet("font-weight: bold; font-size: 16px; padding: 10px;")

        self.log = QTextEdit()
        self.log.setReadOnly(True)

        # Adiciona widgets à interface
        self.layout.addWidget(self.label_planilha)
        self.layout.addWidget(self.btn_planilha)
        self.layout.addWidget(self.label_origem)
        self.layout.addWidget(self.btn_origem)
        self.layout.addWidget(self.label_destino)
        self.layout.addWidget(self.btn_destino)
        self.layout.addWidget(self.btn_processar)
        self.layout.addWidget(QLabel("Log de processamento:"))
        self.layout.addWidget(self.log)

        self.setLayout(self.layout)

    def selecionar_planilha(self):
        path, _ = QFileDialog.getOpenFileName(self, "Selecionar Planilha", "", "Planilhas (*.csv *.xlsx)")
        if path:
            try:
                ext = os.path.splitext(path)[1].lower()
                if ext == ".csv":
                    df = pd.read_csv(path)
                else:
                    df = pd.read_excel(path)

                if 'Categoria' not in df.columns:
                    QMessageBox.critical(self, "Erro", "A planilha não possui a coluna 'Categoria'.")
                    return

                categorias_unicas = df['Categoria'].astype(str).str.strip().unique().tolist()
                selector = CategoriaSelector(categorias_unicas, self)
                if selector.exec_():
                    self.categorias_selecionadas = selector.selecionadas
                    self.planilha_path = path
                    self.label_planilha.setText(f"Planilha: {path}")
                else:
                    self.log.append("❌ Nenhuma categoria selecionada. Processo cancelado.")

            except Exception as e:
                QMessageBox.critical(self, "Erro", f"Erro ao ler a planilha: {e}")

    def selecionar_origem(self):
        path = QFileDialog.getExistingDirectory(self, "Selecionar Pasta de Origem")
        if path:
            self.origem_dir = path
            self.label_origem.setText(f"Pasta de origem: {path}")

    def selecionar_destino(self):
        path = QFileDialog.getExistingDirectory(self, "Selecionar Pasta de Destino")
        if path:
            self.destino_dir = path
            self.label_destino.setText(f"Destino: {path}")

    def log_append(self, text):
        self.log.append(text)
        self.log.verticalScrollBar().setValue(self.log.verticalScrollBar().maximum())

    def processar(self):
        if not all([self.planilha_path, self.origem_dir, self.destino_dir]):
            self.log_append("❌ Por favor, selecione todos os caminhos antes de processar.")
            return

        try:
            ext = os.path.splitext(self.planilha_path)[1].lower()
            if ext == ".csv":
                df = pd.read_csv(self.planilha_path)
            else:
                df = pd.read_excel(self.planilha_path)

            # Aplica filtro por categoria
            df = df[df['Categoria'].astype(str).str.strip().isin(self.categorias_selecionadas)]

        except Exception as e:
            self.log_append(f"❌ Erro ao ler a planilha: {e}")
            return

        colunas_obrigatorias = ['Bloco', 'Unidade', 'Número da Operação', 'Número da OS']
        if not all(col in df.columns for col in colunas_obrigatorias):
            self.log_append(f"⚠️ A planilha precisa conter as colunas: {', '.join(colunas_obrigatorias)}")
            return

        data_str = datetime.today().strftime("%d_%m_%Y")
        arquivos = [f for f in os.listdir(self.origem_dir) if os.path.isfile(os.path.join(self.origem_dir, f))]

        if not arquivos:
            self.log_append("⚠️ Nenhum arquivo encontrado na pasta de origem.")
            return

        self.log_append(f"🔁 Iniciando processamento de {len(arquivos)} arquivos...\n")

        for arquivo in arquivos:
            nome_arquivo = os.path.basename(arquivo)
            prefixo = nome_arquivo.split("_")[0].strip()
            match = df[df['Bloco'].astype(str).str.strip() == prefixo]

            if match.empty:
                destino = os.path.join(self.destino_dir, "Sem_Correspondencia")
                os.makedirs(destino, exist_ok=True)
                shutil.copy2(os.path.join(self.origem_dir, arquivo), os.path.join(destino, nome_arquivo))
                self.log_append(f"🔸 {nome_arquivo} → Sem correspondência (movido para '{destino}')")
                continue

            row = match.iloc[-1]
            unidade = str(row['Unidade']).strip()
            op = str(row['Número da Operação']).strip()
            os_num = str(row['Número da OS']).strip()

            destino = os.path.join(self.destino_dir, unidade, data_str, op, os_num)
            os.makedirs(destino, exist_ok=True)

            shutil.copy2(os.path.join(self.origem_dir, arquivo), os.path.join(destino, nome_arquivo))
            self.log_append(f"✅ {nome_arquivo} → {unidade}/{data_str}/{op}/{os_num}")

        self.log_append("\n✅ Processamento concluído com sucesso!")

if __name__ == "__main__":
    app = QApplication([])
    janela = OrganizadorArquivos()
    janela.show()
    app.exec_()
