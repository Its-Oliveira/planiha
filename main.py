import pandas as pd
import func


df = pd.read_excel("orcamento.xlsx", header=None)
cb = pd.read_excel('cab.xlsx',header=None)

l0 = int(input("Linha primeiro item: "))-1
lf = int(input("Linha ultimo item: "))
tamanho_plan = len(df.index)    


for i in range(l0, lf): # aplicando formatação de itens
    df.loc[i, 0] = func.format_itemizacao(df.loc[i, 0])

for i in range(0,l0): #removendo linhas vazias iniciais
   linha = df.loc[i,0]
   linha = i
   df = df.drop(linha)

for i in range(lf,tamanho_plan): #removendo linhas vazias finais
    linha = df.loc[i,0]
    linha = i
    df = df.drop(linha)

func.SyntaxBancos(df,l0,lf) #formatando bancos

func.codes(df,l0,lf)

df_final = pd.concat([cb, df])
df_final.to_excel('Planilha Ajustada.xlsx', index=False, header=False)