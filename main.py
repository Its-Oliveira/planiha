import pandas as pd
import func
import openpyxl 
from openpyxl.styles import Font, Alignment
from openpyxl.styles import Border, Side


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

wb = openpyxl.load_workbook('Planilha Ajustada.xlsx')
ws = wb["Sheet1"]

for col in ws.columns:
    ws.column_dimensions['A'].width = 8
    ws.column_dimensions['B'].width = 12
    ws.column_dimensions['C'].width = 12
    ws.column_dimensions['D'].width = 80
    ws.column_dimensions['E'].width = 8
    ws.column_dimensions['F'].width = 10
    ws.column_dimensions['G'].width = 12
    ws.column_dimensions['H'].width = 12
    ws.column_dimensions['I'].width = 15
ws.row_dimensions[1].height = 30.0

ws = wb.active
fonte_negrito = Font(bold= True, size=11,name='Arial')
alinhado_centro = Alignment(horizontal='left',vertical='center')
alinhado_esquerda = Alignment(horizontal="left",vertical='center')
borda_fina = Side(border_style="thin", color="000000")

for row in ws['A1':"I1"]:
    for cell in row:
        cell.alignment = alinhado_centro
        cell.font = fonte_negrito
        cell.border = Border(left=borda_fina, right=borda_fina, top=borda_fina, bottom=borda_fina)
    

for row in ws.iter_rows():
    for cell in row:
        cell.alignment = alinhado_esquerda
        cell.border = Border(left=borda_fina, right=borda_fina, top=borda_fina, bottom=borda_fina)


wb.save('Planilha Ajustada.xlsx')