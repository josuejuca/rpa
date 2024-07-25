import sys
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QLineEdit, QComboBox, QPushButton, QVBoxLayout, QFormLayout, QHBoxLayout, QGroupBox, QFrame, QMessageBox, QFileDialog
from PyQt5.QtGui import QIcon, QFont
from PyQt5.QtCore import Qt
import pyautogui as rpa
import time
import pandas as pd

class CertidaoApp(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        # Configurações da janela
        self.setWindowTitle('Solicitação de Certidão SEFAZ')
        self.setGeometry(100, 100, 450, 400)
        self.setWindowIcon(QIcon('icon.png'))

        # Layout principal
        main_layout = QVBoxLayout()
        
        # Seção Propriedades do Empreendimento
        empreendimento_group = QGroupBox('Propriedades do Empreendimento')
        empreendimento_layout = QFormLayout()

        self.nome_empreendimento_input = QLineEdit(self)
        self.nome_empreendimento_input.setPlaceholderText('Nome do empreendimento')
        empreendimento_layout.addRow('Nome do Empreendimento:', self.nome_empreendimento_input)

        self.nome_proprietario_input = QLineEdit(self)
        self.nome_proprietario_input.setPlaceholderText('Nome do proprietário')
        empreendimento_layout.addRow('Nome do Proprietário:', self.nome_proprietario_input)

        empreendimento_group.setLayout(empreendimento_layout)
        main_layout.addWidget(empreendimento_group)

        # Linha divisória
        line1 = QFrame()
        line1.setFrameShape(QFrame.HLine)
        line1.setFrameShadow(QFrame.Sunken)
        main_layout.addWidget(line1)

        # Seção Propriedades do Documento
        documento_group = QGroupBox('Propriedades do Documento')
        documento_layout = QFormLayout()

        self.certidao_combo = QComboBox(self)
        self.certidao_combo.addItems(['Pessoa Fisica', 'Pessoa Jurídica', 'Imóvel'])
        self.certidao_combo.currentIndexChanged.connect(self.updateInputPlaceholder)
        documento_layout.addRow('Tipo de Certidão:', self.certidao_combo)

        self.cpf_cnpj_iptu_input = QLineEdit(self)
        self.updateInputPlaceholder()  # Atualiza o placeholder inicial
        documento_layout.addRow('Documento:', self.cpf_cnpj_iptu_input)

        documento_group.setLayout(documento_layout)
        main_layout.addWidget(documento_group)

        # Linha divisória
        line2 = QFrame()
        line2.setFrameShape(QFrame.HLine)
        line2.setFrameShadow(QFrame.Sunken)
        main_layout.addWidget(line2)

        # Botão para selecionar arquivo Excel
        self.select_file_button = QPushButton('Selecionar Planilha Excel', self)
        self.select_file_button.setFont(QFont('Arial', 12))
        self.select_file_button.setStyleSheet('background-color: #2196F3; color: white; padding: 10px;')
        self.select_file_button.clicked.connect(self.loadExcelFile)

        # Adiciona o botão ao layout principal
        file_button_layout = QHBoxLayout()
        file_button_layout.addStretch(1)
        file_button_layout.addWidget(self.select_file_button)
        file_button_layout.addStretch(1)
        main_layout.addLayout(file_button_layout)

        # Botão de Submissão
        self.submit_button = QPushButton('Solicitar Certidão', self)
        self.submit_button.setFont(QFont('Arial', 12))
        self.submit_button.setStyleSheet('background-color: #4CAF50; color: white; padding: 10px;')
        self.submit_button.clicked.connect(self.onSubmit)

        # Adiciona o botão ao layout principal
        button_layout = QHBoxLayout()
        button_layout.addStretch(1)
        button_layout.addWidget(self.submit_button)
        button_layout.addStretch(1)
        main_layout.addLayout(button_layout)

        # Configura o layout principal
        self.setLayout(main_layout)

    def updateInputPlaceholder(self):
        tipo_certidao = self.certidao_combo.currentText()
        if tipo_certidao == 'Pessoa Fisica':
            self.cpf_cnpj_iptu_input.setPlaceholderText('Digite o CPF')
            self.cpf_cnpj_iptu_input.setMaxLength(14)  # Formato: 000.000.000-00
        elif tipo_certidao == 'Pessoa Jurídica':
            self.cpf_cnpj_iptu_input.setPlaceholderText('Digite o CNPJ')
            self.cpf_cnpj_iptu_input.setMaxLength(18)  # Formato: 00.000.000/0000-00
        else:
            self.cpf_cnpj_iptu_input.setPlaceholderText('Digite a Inscrição do IPTU')
            self.cpf_cnpj_iptu_input.setMaxLength(30)  # Pode conter letras e números

    def loadExcelFile(self):
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getOpenFileName(self, "Selecionar Planilha Excel", "", "Arquivos Excel (*.xlsx);;Todos os Arquivos (*)", options=options)
        if file_name:
            self.df = pd.read_excel(file_name)
            QMessageBox.information(self, 'Sucesso', f'Planilha carregada com {len(self.df)} registros')

    def onSubmit(self):
        # Na planilha a primeira linha aonde fica os CPFs tem que ser "Documento"
        nome_empreendimento = self.nome_empreendimento_input.text()
        nome_proprietario = self.nome_proprietario_input.text()
        documento = self.cpf_cnpj_iptu_input.text()
        tipo_certidao = self.certidao_combo.currentText()

        if hasattr(self, 'df'):
            rpa.alert(text="O BOT vai começar", title="Início")
            for i, doc in enumerate(self.df['Documento']):
                doc = str(doc).zfill(11)
                if not self.validateDocument(doc, tipo_certidao):
                    QMessageBox.warning(self, 'Erro', f'Documento inválido: {doc}')
                else:
                    is_last = (i == len(self.df['Documento']) - 1)
                    self.startAutomation(doc, is_last)
        elif documento:
            if not self.validateDocument(documento, tipo_certidao):
                QMessageBox.warning(self, 'Erro', 'Documento inválido!')
            else:
                rpa.alert(text="O BOT vai começar", title="Início")
                self.startAutomation(documento, True)
        else:
            QMessageBox.warning(self, 'Erro', 'Nenhum documento fornecido!')

    def validateDocument(self, doc, tipo_certidao):
        if tipo_certidao == 'Pessoa Fisica':
            return self.validateCPF(doc)
        elif tipo_certidao == 'Pessoa Jurídica':
            return self.validateCNPJ(doc)
        else:
            return True  # Validação genérica para Inscrição do IPTU

    def validateCPF(self, cpf):
        # Implementação básica de validação de CPF
        cpf = ''.join(filter(str.isdigit, cpf))
        if len(cpf) != 11:
            return False
        if cpf in [s * 11 for s in "0123456789"]:
            return False
        return True

    def validateCNPJ(self, cnpj):
        # Implementação básica de validação de CNPJ
        cnpj = ''.join(filter(str.isdigit, cnpj))
        if len(cnpj) != 14:
            return False
        if cnpj in [s * 14 for s in "0123456789"]:
            return False
        return True

    def startAutomation(self, doc, is_last):
        rpa.PAUSE = 0.5
        
        self.open_browser_and_navigate("https://ww1.receita.fazenda.df.gov.br/cidadao/certidoes/Certidao")
        
        self.select_certificate_type(doc)
        # self.select_region()
        # self.enter_document(doc)
        self.baixa_certidao(doc, is_last)
        
    def open_browser_and_navigate(self, url):
        """Abre o navegador e navega para a URL especificada."""
        time.sleep(1)
        rpa.press('win')
        time.sleep(0.5)
        rpa.write("chrome")
        rpa.press('enter')
        time.sleep(2)
        rpa.write(url)
        time.sleep(1)
        rpa.press('enter')

    def select_certificate_type(self, doc):
        """Seleciona o tipo de certidão no site."""
        tipo_certidao = self.certidao_combo.currentText()
        if tipo_certidao == 'Pessoa Fisica':
            time.sleep(1)
            rpa.moveTo(356, 428)  # Ajustar de acordo com a tela 
            time.sleep(0.5)
            rpa.click()
            # seleciona o tipo de documento 
            time.sleep(1)
            rpa.press('tab')
            time.sleep(0.5)
            rpa.press('tab')
            time.sleep(0.5)
            rpa.write(doc, interval=0.1)
            for _ in range(4):
                time.sleep(0.5)
                rpa.press('tab')
            time.sleep(1)
            rpa.press('space')
            time.sleep(2)

        elif tipo_certidao == 'Pessoa Jurídica':
            time.sleep(1)
            rpa.moveTo(364, 448)  # Ajustar de acordo com a tela 
            time.sleep(0.5)
            rpa.click()
            time.sleep(1)
            rpa.press('tab')
            time.sleep(0.5)
            rpa.press('right')
            time.sleep(0.5)
            rpa.press('tab')
            time.sleep(0.5)
            rpa.write(doc, interval=0.1)
            for _ in range(4):
                time.sleep(0.5)
                rpa.press('tab')
            time.sleep(1)
            rpa.press('space')
            time.sleep(2)
        elif tipo_certidao == 'Imóvel':
            time.sleep(1)
            rpa.moveTo(364, 448)  # Ajustar de acordo com a tela 
            time.sleep(0.5)
            rpa.click()
            time.sleep(1)
            rpa.press('tab')
            time.sleep(0.5)
            rpa.press('right')
            time.sleep(0.5)
            rpa.press('right')
            time.sleep(0.5)
            rpa.press('tab')
            time.sleep(0.5)
            rpa.write(doc, interval=0.1)
            for _ in range(4):
                time.sleep(0.5)
                rpa.press('tab')
            time.sleep(1)
            rpa.press('space')
            time.sleep(2)

    def baixa_certidao(self, cpf, is_last):
        """Baixa e salva o arquivo."""
        time.sleep(1)
        rpa.press('enter')
        tipo_certidao = self.certidao_combo.currentText()
        
        if tipo_certidao == 'Pessoa Fisica':
            rpa.write(cpf + '-sefaz-pf', interval=0.1)
            rpa.press('enter')
        elif tipo_certidao == 'Pessoa Jurídica':
            rpa.write(cpf + '-sefaz-pj', interval=0.1)
            rpa.press('enter')
        else:
            rpa.write(cpf + '-sefaz-imovel', interval=0.1)
            rpa.press('enter')
            
        if is_last:
            rpa.alert(text="Processo finalizado, todas as certidões foram salvas!", title="Processo finalizado")


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = CertidaoApp()
    ex.show()
    sys.exit(app.exec_())
