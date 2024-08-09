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

        # Botão para iniciar o processo
        self.start_button = QPushButton('Iniciar', self)
        self.start_button.clicked.connect(self.processar_certidoes)
        main_layout.addWidget(self.start_button)

        self.setLayout(main_layout)

    def sefaz(self):
        # Código para obter certidão SEFaz
        print("Obtendo certidão SEFaz...")
        # Simulação de espera
        time.sleep(2)
        print("Certidão SEFaz obtida.")

    def receita(self):
        # Código para obter certidão Receita
        print("Obtendo certidão Receita...")
        # Simulação de espera
        time.sleep(2)
        print("Certidão Receita obtida.")

    def falencia(self):
        # Código para obter certidão Falência
        print("Obtendo certidão Falência...")
        # Simulação de espera
        time.sleep(2)
        print("Certidão Falência obtida.")

    def trabalhista(self):
        # Código para obter certidão Trabalhista
        print("Obtendo certidão Trabalhista...")
        # Simulação de espera
        time.sleep(2)
        print("Certidão Trabalhista obtida.")

    def civel(self, cpf, url_trf1):
        # "Cível", "Criminal", "Eleitoral"
        """Abre o navegador e navega para a URL especificada."""        
        # rpa.press('winleft')
        time.sleep(1)
        rpa.press('win')
        time.sleep(0.5)
        rpa.write("chrome")
        rpa.press('enter')
        time.sleep(2)
        rpa.write(url_trf1)
        time.sleep(1)
        rpa.press('enter')
        """Seleciona o tipo de certidão no site."""
        time.sleep(1)
        rpa.press('space')        
        rpa.press('down')  # escolhe o tipo de certidão       
        rpa.press('space')
        """Seleciona a região desejada."""
        rpa.press('tab')
        time.sleep(0.5)
        rpa.press('space')
        for _ in range(5):
            rpa.press('down')
        time.sleep(1)
        rpa.press('enter')
        """Entra com o CPF fornecido pelo usuário."""
        time.sleep(1)
        rpa.press('tab')
        time.sleep(0.5)
        rpa.press('tab')
        time.sleep(0.5)
        rpa.write(cpf, interval=0.1)
        time.sleep(3)
        rpa.press('enter')
        """Baixa e salva o arquivo."""
        time.sleep(5)
        rpa.moveTo(684, 197)  # Ajustar de acordo com a tela 
        rpa.sleep(0.5)
        rpa.click()
        time.sleep(0.5)
        # baixar arquivo ( x 1260 y 114)
        rpa.moveTo(1260, 114)  # Ajustar de acordo com a tela 
        rpa.sleep(0.5)
        rpa.click()
        # salvar arquivo 
        time.sleep(1)
        rpa.write(cpf + '-civel', interval=0.1)
        rpa.press('enter')

    def criminal(self, cpf, url_trf1):
        # "Cível", "Criminal", "Eleitoral"
        """Abre o navegador e navega para a URL especificada."""        
        # rpa.press('winleft')
        time.sleep(1)
        rpa.press('win')
        time.sleep(0.5)
        rpa.write("chrome")
        rpa.press('enter')
        time.sleep(2)
        rpa.write(url_trf1)
        time.sleep(1)
        rpa.press('enter')
        """Seleciona o tipo de certidão no site."""
        time.sleep(1)
        rpa.press('space')       
        rpa.press('down') # Escolhe a certidão 
        time.sleep(0.5)
        rpa.press('down') # Escolhe a certidão        
        rpa.press('space')
        """Seleciona a região desejada."""
        rpa.press('tab')
        time.sleep(0.5)
        rpa.press('space')
        for _ in range(5):
            rpa.press('down')
        time.sleep(1)
        rpa.press('enter')
        """Entra com o CPF fornecido pelo usuário."""
        time.sleep(1)
        rpa.press('tab')
        time.sleep(0.5)
        rpa.press('tab')
        time.sleep(0.5)
        rpa.write(cpf, interval=0.1)
        time.sleep(3)
        rpa.press('enter')
        """Baixa e salva o arquivo."""
        time.sleep(5)
        rpa.moveTo(684, 197)  # Ajustar de acordo com a tela 
        rpa.sleep(0.5)
        rpa.click()
        time.sleep(0.5)
        # baixar arquivo ( x 1260 y 114)
        rpa.moveTo(1260, 114)  # Ajustar de acordo com a tela 
        rpa.sleep(0.5)
        rpa.click()
        # salvar arquivo 
        time.sleep(1)
        rpa.write(cpf + '-criminal', interval=0.1)
        rpa.press('enter')

    def eleitoral(self, cpf, url_trf1):
        # "Cível", "Criminal", "Eleitoral"
        """Abre o navegador e navega para a URL especificada."""        
        # rpa.press('winleft')
        time.sleep(1)
        rpa.press('win')
        time.sleep(0.5)
        rpa.write("chrome")
        rpa.press('enter')
        time.sleep(2)
        rpa.write(url_trf1)
        time.sleep(1)
        rpa.press('enter')
        """Seleciona o tipo de certidão no site."""
        time.sleep(1)
        rpa.press('space')       
        rpa.press('down') # Escolhe a certidão 
        time.sleep(0.5)
        rpa.press('down') # Escolhe a certidão 
        time.sleep(0.5)
        rpa.press('down') # Escolhe a certidão 
        rpa.press('space')
        """Seleciona a região desejada."""
        rpa.press('tab')
        time.sleep(0.5)
        rpa.press('space')
        for _ in range(5):
            rpa.press('down')
        time.sleep(1)
        rpa.press('enter')
        """Entra com o CPF fornecido pelo usuário."""
        time.sleep(1)
        rpa.press('tab')
        time.sleep(0.5)
        rpa.press('tab')
        time.sleep(0.5)
        rpa.write(cpf, interval=0.1)
        time.sleep(3)
        rpa.press('enter')
        """Baixa e salva o arquivo."""
        time.sleep(5)
        rpa.moveTo(684, 197)  # Ajustar de acordo com a tela 
        rpa.sleep(0.5)
        rpa.click()
        time.sleep(0.5)
        # baixar arquivo ( x 1260 y 114)
        rpa.moveTo(1260, 114)  # Ajustar de acordo com a tela 
        rpa.sleep(0.5)
        rpa.click()
        # salvar arquivo 
        time.sleep(1)
        rpa.write(cpf + '-eleitoral', interval=0.1)
        rpa.press('enter')

    def processar_certidoes(self):
        cpf = self.cpf_input.text()
        url_trf1 = 'https://sistemas.trf1.jus.br/certidao/#/solicitacao'
        rpa.alert(text="O BOT vai começar", title="Início")
        self.civel(cpf, url_trf1)
        self.criminal(cpf, url_trf1)
        self.eleitoral(cpf, url_trf1 )
        self.sefaz()
        self.receita()
        self.falencia()
        self.trabalhista()
        QMessageBox.information(self, 'Concluído', 'Todas as certidões foram obtidas com sucesso.')

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = CertidaoApp()
    ex.show()
    sys.exit(app.exec_())
