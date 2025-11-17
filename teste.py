import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import func
import os

def processar():
    # 1) validação dos inputs
    primeira = entrada_primeira.get().strip()
    ultima = entrada_ultima.get().strip()

    if primeira == "" or ultima == "":
        messagebox.showerror("Erro", "Preencha as duas caixas: Primeira linha e Última linha.")
        return

    if not (primeira.isdigit() and ultima.isdigit()):
        messagebox.showerror("Erro", "Digite apenas números inteiros nas caixas de linha.")
        return

    l0 = int(primeira) - 1
    lf = int(ultima) - 1

    if l0 < 0 or lf < 0:
        messagebox.showerror("Erro", "Os números das linhas devem ser >= 1.")
        return
    if lf < l0:
        messagebox.showerror("Erro", "A última linha deve ser maior ou igual à primeira.")
        return

    caminho_orc = filedialog.askopenfilename(
        title="Selecione o arquivo orcamento.xlsx",
        filetypes=[("Arquivos Excel", "*.xlsx *.xls")]
    )
    if not caminho_orc:
        print("Operação cancelada: nenhum arquivo selecionado.")
        return

    # carregar arquivos e checar limites
    try:
        df = pd.read_excel(caminho_orc, header=None)
    except Exception as e:
        messagebox.showerror("Erro ao ler Excel", f"Não foi possível ler o arquivo selecionado:\n{e}")
        return

    # carregar cabeçalho
    cab_path = "cab.xlsx"
    if not os.path.exists(cab_path):
        messagebox.showerror("Erro", f"Arquivo de cabeçalho '{cab_path}' não encontrado.")
        return
    try:
        cb = pd.read_excel(cab_path, header=None)
    except Exception as e:
        messagebox.showerror("Erro ao ler cab.xlsx", f"Não foi possível ler 'cab.xlsx':\n{e}")
        return

    tamanho_plan = len(df.index)
    if l0 >= tamanho_plan or lf >= tamanho_plan:
        messagebox.showerror("Erro", f"Os números das linhas estão fora do intervalo do arquivo.\nO arquivo tem {tamanho_plan} linhas (índices 1..{tamanho_plan}).")
        return

    try:
        # 4) recortar o bloco desejado (inclui lf)
        df_slice = df.iloc[l0:lf+1].copy()  # .copy() para evitar avisos e trabalhar seguro
        df_slice = df_slice.reset_index(drop=True)  # reindexa para 0..n-1

        # 5) aplicar format_itemizacao no intervalo (agora 0..len-1)
        for i in range(len(df_slice)):
            # usa .at por ser mais rápido para uma célula
            df_slice.at[i, 0] = func.format_itemizacao(df_slice.at[i, 0])

        # 6) chamar funções auxiliares: como reindexamos, passamos o intervalo relativo
        l0_rel = 0
        lf_rel = len(df_slice) - 1

        # Se suas funções precisam dos índices originais (antes do slice),
        # substitua acima por l0_original e lf_original conforme necessário.
        func.SyntaxBancos(df_slice, l0_rel, lf_rel)
        func.codes(df_slice, l0_rel, lf_rel)

        # 7) concatenar cabeçalho fixo com o bloco processado e salvar
        df_final = pd.concat([cb, df_slice], ignore_index=True)
        saida = 'Planilha Ajustada.xlsx'
        df_final.to_excel(saida, index=False, header=False)

    except Exception as e:
        # mostra a exceção concreta para diagnóstico
        messagebox.showerror("Erro durante o processamento", f"Ocorreu um erro:\n{e}")
        import traceback
        print("TRACEBACK (console):")
        traceback.print_exc()
        return

    messagebox.showinfo("Concluído", f"✅ Processamento concluído!\nArquivo gerado: {saida}")
    print(f"Processamento concluído. Arquivo salvo em: {saida}")


# ----------- Interface -----------
janela = tk.Tk()
janela.title("Processar Excel")
janela.geometry("380x220")

label1 = tk.Label(janela, text="Primeira linha (1 = primeira linha do arquivo):")
label1.pack(anchor='w', padx=10, pady=(10,0))
entrada_primeira = tk.Entry(janela, width=20)
entrada_primeira.pack(padx=10, pady=5)

label2 = tk.Label(janela, text="Última linha (inclusiva):")
label2.pack(anchor='w', padx=10, pady=(5,0))
entrada_ultima = tk.Entry(janela, width=20)
entrada_ultima.pack(padx=10, pady=5)

botao = tk.Button(janela, text="Selecionar Arquivo e Processar", command=processar)
botao.pack(pady=15)

janela.mainloop()