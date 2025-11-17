import pandas as pd
import tkinter as tk
from tkinter import filedialog

def selecionar_arquivo():

    tipoArquivo = [('Arquivos excel', '*.xlsx *.xls')]
    caminhoArquivo = filedialog.askopenfilename(filetypes=tipoArquivo)

    if caminhoArquivo:
        try: 
            df = pd.read_excel(caminhoArquivo)
            print("Arquivo carregado com sucesso")
            print(df.head())
        except Exception as e:
            print("Erro ao carregar arquivo {e}")

janela = tk.Tk()
janela.title("Selecionar arquivo XLSX")
janela.geometry("300x150")


botao_selecionar = tk.Button(janela, text="Selecionar Arquivo XLSX", command=selecionar_arquivo)
botao_selecionar.pack(pady=30, padx=30)

janela.mainloop()