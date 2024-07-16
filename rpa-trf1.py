import pyautogui as rpa
import tkinter as tk
from tkinter import simpledialog

# URL do site para solicitar a certidão
URL = "https://sistemas.trf1.jus.br/certidao/#/solicitacao"

def get_input():
    """Obtém a entrada do usuário usando uma caixa de diálogo Tkinter."""
    root = tk.Tk()
    root.withdraw()  # Esconde a janela principal
    user_input = simpledialog.askstring(title="Input", prompt="Digite o CPF:")
    root.destroy()
    return user_input

def open_browser_and_navigate(url):
    """Abre o navegador e navega para a URL especificada."""
    rpa.press('winleft')
    rpa.write("chrome")
    rpa.press('enter')
    rpa.sleep(2)
    rpa.write(url)
    rpa.sleep(1)
    rpa.press('enter')

def select_certificate_type():
    """Seleciona o tipo de certidão no site."""
    rpa.sleep(1)
    rpa.press('space')
    rpa.press('down')
    rpa.press('space')

def select_region():
    """Seleciona a região desejada."""
    rpa.press('tab')
    rpa.sleep(0.5)
    rpa.press('space')
    for _ in range(5):
        rpa.press('down')
    rpa.sleep(1)
    rpa.press('enter')

def enter_cpf(cpf):
    """Entra com o CPF fornecido pelo usuário."""
    rpa.sleep(1)
    rpa.press('tab')
    rpa.sleep(0.5)
    rpa.press('tab')
    rpa.sleep(0.5)
    rpa.write(cpf, interval=0.1)
    rpa.sleep(3)
    rpa.press('enter')

def main():
    cpf = get_input()
    
    if cpf:
        rpa.alert(text="O BOT vai começar", title="Início")
        rpa.PAUSE = 0.5
        
        open_browser_and_navigate(URL)
        
        select_certificate_type()
        
        select_region()
        
        enter_cpf(cpf)
    else:
        rpa.alert(text='Não foi informado o CPF', title='Erro')

if __name__ == "__main__":
    main()
