import time
from selenium import webdriver
from tkinter import Tk, Button, Label, filedialog
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

caminho_arquivo = None

def selecionar_arquivo():
    global caminho_arquivo
    caminho_arquivo = filedialog.askopenfilename()

    if caminho_arquivo:
        label_arquivo.config(text=f"Arquivo: {caminho_arquivo}")

        botao_selecionar.pack_forget()

        botao_iniciar.pack(pady=10)
    else:
        label_arquivo.config(text="Nenhum arquivo selecionado!")

    return caminho_arquivo


def criar_interface():
    global root
    root = Tk()
    root.title("WhatsAuto")
    root.geometry("400x150")

    global label_arquivo
    label_arquivo = Label(root, text="Selecione a planilha de dados para iniciar!", wraplength=300)
    label_arquivo.pack(pady=20)

    global botao_selecionar
    botao_selecionar = Button(root, text="Selecione a planilha", command=selecionar_arquivo)
    botao_selecionar.pack(pady=10)

    global botao_iniciar
    botao_iniciar = Button(root, text="Clique aqui para iniciar o processo", command=iniciar)

    root.mainloop()


def iniciar():
    root.destroy()
    browser = webdriver.Chrome()
    browser.get("https://web.whatsapp.com/")

    try:
        xpath_search = '//*[@id="side"]/div[1]/div/div[2]/div[2]/div/div/p'
        search_element = WebDriverWait(browser, 120).until(
            EC.element_to_be_clickable((By.XPATH, xpath_search))
        )

        search_element.click()

        search_element.send_keys()

    except Exception as e:
        print(f'Erro: {e}')

    time.sleep(5)
    print(caminho_arquivo)


criar_interface()