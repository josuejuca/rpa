import sys
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QLineEdit, QComboBox, QPushButton, QVBoxLayout, QFormLayout, QHBoxLayout, QGroupBox, QFrame, QMessageBox
from PyQt5.QtGui import QIcon, QFont
from PyQt5.QtCore import Qt
import pyautogui as rpa
import time

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

        self.certidao_combo = QComboBox(self)
        self.certidao_combo.addItems(['Cível', 'Criminal', 'Eleitoral'])
        documento_layout.addRow('Tipo de Certidão:', self.certidao_combo)

        documento_group.setLayout(documento_layout)
        main_layout.addWidget(documento_group)

        # Linha divisória
        line2 = QFrame()
        line2.setFrameShape(QFrame.HLine)
        line2.setFrameShadow(QFrame.Sunken)
        main_layout.addWidget(line2)

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

    def onSubmit(self):
        nome_empreendimento = self.nome_empreendimento_input.text()
        nome_proprietario = self.nome_proprietario_input.text()
        cpf = self.cpf_input.text()
        tipo_certidao = self.certidao_combo.currentText()

        if not self.validateCPF(cpf):
            QMessageBox.warning(self, 'Erro', 'CPF inválido!')
        else:
            # QMessageBox.information(self, 'Solicitação Recebida',
            #                         f'Nome do Empreendimento: {nome_empreendimento}\n'
            #                         f'Nome do Proprietário: {nome_proprietario}\n'
            #                         f'CPF: {cpf}\n'
            #                         f'Tipo de Certidão: {tipo_certidao}')
            self.startAutomation(cpf)

    def validateCPF(self, cpf):
        # Implementação básica de validação de CPF
        cpf = ''.join(filter(str.isdigit, cpf))
        if len(cpf) != 11:
            return False
        if cpf in [s * 11 for s in "0123456789"]:
            return False
        return True

    def startAutomation(self, cpf):
        rpa.alert(text="O BOT vai começar", title="Início")
        rpa.PAUSE = 0.5
        
        self.open_browser_and_navigate("https://sistemas.trf1.jus.br/certidao/#/solicitacao")
        
        self.select_certificate_type()
        self.select_region()
        self.enter_cpf(cpf)
        self.baixa_certidao()
        
    def open_browser_and_navigate(self, url):
        """Abre o navegador e navega para a URL especificada."""
        # rpa.press('winleft')
        rpa.press('win')
        time.sleep(0.5)
        rpa.write("chrome")
        rpa.press('enter')
        time.sleep(2)
        rpa.write(url)
        time.sleep(1)
        rpa.press('enter')

    def select_certificate_type(self):
        """Seleciona o tipo de certidão no site."""
        tipo_certidao = self.certidao_combo.currentText()
        
        if ( tipo_certidao == 'Cível'):
            time.sleep(1)
            rpa.press('space')
            rpa.press('down')
            rpa.press('space')
        elif (tipo_certidao == 'Criminal'):
            time.sleep(1)
            rpa.press('space')
            rpa.press('down')
            time.sleep(0.5)
            rpa.press('down')
            rpa.press('space')
        else:
            time.sleep(1)
            rpa.press('space')
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
        
    def baixa_certidao(self, cpf):
        time.sleep(1)
        rpa.press('tab')
        time.sleep(0.5)
        rpa.press('enter')
        for _ in range(8):
            rpa.press('tab')
        time.sleep(1)
        rpa.press('space')
        
        # Salva o doc 
        # C:\Users\Juca\Desktop\qb
        for _ in range(5):
            rpa.press('tab')
        time.sleep(1)
        rpa.press('space')
        time.sleep(0.5)
        rpa.press('enter')
        

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = CertidaoApp()
    ex.show()
    sys.exit(app.exec_())
