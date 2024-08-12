import sys
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QLineEdit, QPushButton, QVBoxLayout, QFormLayout, QHBoxLayout, QGroupBox, QFrame, QMessageBox, QFileDialog
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
        self.setWindowTitle('Solicitação de Certidão')
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

        self.cpf_input = QLineEdit(self)
        self.cpf_input.setPlaceholderText('Digite o CPF')
        self.cpf_input.setMaxLength(14)  # Formato: 000.000.000-00
        documento_layout.addRow('CPF:', self.cpf_input)

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
        self.submit_button = QPushButton('Solicitar Certidões', self)
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

    def loadExcelFile(self):
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getOpenFileName(self, "Selecionar Planilha Excel", "", "Arquivos Excel (*.xlsx);;Todos os Arquivos (*)", options=options)
        if file_name:
            self.df = pd.read_excel(file_name)
            QMessageBox.information(self, 'Sucesso', f'Planilha carregada com {len(self.df)} registros')

    def onSubmit(self):
        nome_empreendimento = self.nome_empreendimento_input.text()
        nome_proprietario = self.nome_proprietario_input.text()
        cpf = self.cpf_input.text()

        if hasattr(self, 'df'):
            rpa.alert(text="O BOT vai começar", title="Início")
            for i, cpf in enumerate(self.df['CPF']):
                cpf = str(cpf).zfill(11)
                if not self.validateCPF(cpf):
                    QMessageBox.warning(self, 'Erro', f'CPF inválido: {cpf}')
                else:
                    is_last = (i == len(self.df['CPF']) - 1)
                    self.startAutomation(cpf, is_last)
        elif cpf:
            if not self.validateCPF(cpf):
                QMessageBox.warning(self, 'Erro', 'CPF inválido!')
            else:
                rpa.alert(text="O BOT vai começar", title="Início")
                self.startAutomation(cpf, True)
        else:
            QMessageBox.warning(self, 'Erro', 'Nenhum CPF fornecido!')

    def validateCPF(self, cpf):
        # Implementação básica de validação de CPF
        cpf = ''.join(filter(str.isdigit, cpf))
        if len(cpf) != 11:
            return False
        if cpf in [s * 11 for s in "0123456789"]:
            return False
        return True

    def startAutomation(self, cpf, is_last):
        rpa.PAUSE = 0.5
        
        certidoes = ["Civel", "Criminal", "Eleitoral"]
        for certidao in certidoes:
            self.open_browser_and_navigate("https://sistemas.trf1.jus.br/certidao/#/solicitacao")
            self.get_certidao(cpf, certidao)
        
        if is_last:
            rpa.alert(text="Processo finalizado, todas as certidões foram salvas!", title="Processo finalizado")

    def open_browser_and_navigate(self, url):
        """Abre o navegador e navega para a URL especificada."""
        # rpa.press('winleft')
        time.sleep(1)
        rpa.press('win')
        time.sleep(0.5)
        rpa.write("brave") # chrome
        rpa.press('enter')
        time.sleep(2)
        rpa.write(url)
        time.sleep(1)
        rpa.press('enter')

    def get_certidao(self, cpf, tipo_certidao):
        """Obtém uma certidão do tipo especificado."""
        self.select_certificate_type(tipo_certidao)
        self.select_region()
        self.enter_cpf(cpf)
        self.baixa_certidao(cpf, tipo_certidao)
        
    def select_certificate_type(self, tipo_certidao):
        """Seleciona o tipo de certidão no site."""
        time.sleep(1)
        rpa.press('space')
        time.sleep(2)
        if tipo_certidao == 'Civel':
            time.sleep(2)
            rpa.press('down')
            time.sleep(1)
        elif tipo_certidao == 'Criminal':
            time.sleep(1)
            rpa.press('down')
            time.sleep(0.5)
            rpa.press('down')
        else:
            time.sleep(1)
            rpa.press('down')
            time.sleep(0.5)
            rpa.press('down')
            time.sleep(0.5)
            rpa.press('down')
        rpa.press('space')

    def select_region(self):
        """Seleciona a região desejada."""
        rpa.press('tab')
        time.sleep(0.5)
        rpa.press('space')
        for _ in range(5):
            rpa.press('down')
        time.sleep(1)
        rpa.press('enter')

    def enter_cpf(self, cpf):
        """Entra com o CPF fornecido pelo usuário."""
        time.sleep(1)
        rpa.press('tab')
        time.sleep(0.5)
        rpa.press('tab')
        time.sleep(0.5)
        rpa.write(cpf, interval=0.1)
        time.sleep(3)
        rpa.press('enter')
        
    def baixa_certidao(self, cpf, tipo_certidao):
        """Baixa e salva o arquivo."""
        time.sleep(5)
        rpa.moveTo(684, 197)  # Ajustar de acordo com a tela 
        rpa.sleep(0.5)
        rpa.click()
        time.sleep(0.5)
        # baixar arquivo ( x 1260 y 114)
        rpa.moveTo(1260, 114)  # Ajustar de acordo com a tela 
        rpa.sleep(1)
        rpa.click()
        # salvar arquivo 
        time.sleep(1)
        rpa.write(cpf + '-' + tipo_certidao.lower(), interval=0.1)
        rpa.press('enter')
        
if __name__ == '__main__':
    app = QApplication(sys.argv)
    certidao_app = CertidaoApp()
    certidao_app.show()
    sys.exit(app.exec_())
