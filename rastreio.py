import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from time import sleep
from tkinter import Tk, Button, Label, Entry, filedialog
import requests
from io import BytesIO
import os

def abrir_selecionador_arquivo():
    arquivo_excel = filedialog.askopenfilename(title="Selecione o arquivo Excel com os códigos de rastreio", filetypes=[("Excel Files", "*.xlsx")])
    entry_caminho_arquivo.delete(0, 'end')
    entry_caminho_arquivo.insert(0, arquivo_excel)

def processar_rastreio():
    arquivo_excel = entry_caminho_arquivo.get()

    if not arquivo_excel:
        print("Erro: Por favor, selecione um arquivo Excel.")
        return

    if not os.path.exists(arquivo_excel):
        print("Erro: O arquivo Excel selecionado não existe.")
        return

    try:
        df = pd.read_excel(arquivo_excel)

        if 'Codigo_Rastreio' not in df.columns:
            print("Erro: O arquivo Excel não contém uma coluna chamada 'Codigo_Rastreio'.")
            return

        driver = webdriver.Chrome()

        for index, row in df.iterrows():
            cod_rastreio = row['Codigo_Rastreio']

            driver.get(f'http://rastreamento.visualset.com.br/index.php?parametros={cod_rastreio}')
            sleep(3)

            ultima_atualizacao = driver.find_element(By.XPATH, "//*[@id='grid_tabelarastreamento_rec_0']").text

            df.at[index, 'Ultima_Atualizacao'] = ultima_atualizacao

        df.to_excel('rastreio_atualizado.xlsx', index=False)

        print("Processamento concluído com sucesso.")

    except Exception as e:
        print("Erro durante o processamento:", str(e))

    finally:
        try:
            driver.quit()
        except NameError:
            pass

def download_arquivo_exemplo():
    url = 'https://ceopag.com.br/rastreio.xlsx'
    response = requests.get(url)
    arquivo_excel_bytes = BytesIO(response.content)

    diretorio_destino = filedialog.askdirectory()
    if diretorio_destino:  
        caminho_destino = os.path.join(diretorio_destino, 'exemplo_rastreio.xlsx')
        with open(caminho_destino, 'wb') as f:
            f.write(arquivo_excel_bytes.getbuffer())

root = Tk()
root.title("Rastreamento de Códigos")

label_caminho_arquivo = Label(root, text="Caminho do arquivo:")
label_caminho_arquivo.grid(row=0, column=0, padx=5, pady=5)

entry_caminho_arquivo = Entry(root, width=50)
entry_caminho_arquivo.grid(row=0, column=1, padx=5, pady=5)

btn_selecionar_arquivo = Button(root, text="Selecionar Arquivo", command=abrir_selecionador_arquivo)
btn_selecionar_arquivo.grid(row=0, column=2, padx=5, pady=5)

btn_processar_rastreio = Button(root, text="Realizar Rastreios", command=processar_rastreio)
btn_processar_rastreio.grid(row=1, column=0, columnspan=3, padx=5, pady=5)

btn_download_exemplo = Button(root, text="Baixar excel a ser preenchido", command=download_arquivo_exemplo)
btn_download_exemplo.grid(row=2, column=0, columnspan=3, padx=5, pady=5)

root.mainloop()
