import pandas as pd
import func
import openpyxl 

df = pd.read_excel("plan cliente.xlsx", header=None)
cb = pd.read_excel('cab.xlsx',header=None)


l0 = int(input("Linha primeiro item: "))-1
lf = int(input("Linha ultimo item: "))


col_item = str(input('col item: '))
if col_item == 'A' or col_item == 'a':
    df.rename(columns={0 : col_item}, inplace=True)
elif col_item == 'B' or col_item == 'b':
    df.rename(columns={1 : col_item}, inplace=True)
elif col_item == 'C' or col_item == 'c':
    df.rename(columns={2 : col_item}, inplace=True)

col_cod = str(input('col cod: '))
if col_cod == 'A' or col_cod == 'a':
    df.rename(columns={0 : col_cod}, inplace=True)
elif col_cod == 'B' or col_cod == 'b':
    df.rename(columns={1 : col_cod}, inplace=True)
elif col_cod == 'C' or col_cod == 'c':
    df.rename(columns={2 : col_cod}, inplace=True)

col_banc = str(input('col banc: '))
if col_banc == 'A' or col_banc == 'a':
    df.rename(columns={0 : col_banc}, inplace=True)
elif col_banc == 'B' or col_banc == 'b':
    df.rename(columns={1 : col_banc}, inplace=True)
elif col_banc == 'C' or col_banc == 'c':
    df.rename(columns={2 : col_banc}, inplace=True)


df = df.reindex([col_item,col_cod,col_banc,3,4,5,6,7,8], axis=1)

df.rename(columns={col_cod:1}, inplace=True)
df.rename(columns={col_item:0}, inplace=True)
df.rename(columns={col_banc:2}, inplace=True)


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
wb = openpyxl.load_workbook('Planilha Ajustada.xlsx')
ws = wb["Sheet1"]


func.format_plan_ajustada(ws,wb)

