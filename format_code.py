import pandas as pd
import func

df = pd.read_excel("orcamento.xlsx", header=None)
tamanho_plan = len(df.index)

l0 = int(input("Linha primeiro item: "))-1
lf = int(input("Linha ultimo item: "))

func.SyntaxBancos(df,l0,lf)

def codes(df,l0,lf):
    for i in range(l0,lf):
        bc = df.loc[i,2]
        if bc == "SINAPI":
            code = str(df.loc[i,1])
            codel = (code.strip())
            df.loc[i,1] = codel
            if len(codel) == 5 or len(codel) == 6:
                df.loc[i,1] = codel
            else:
                df.loc[i,1] = "Erro, formato não reconhecido pela SINAPI"

codes()


df.to_excel('Planilha Ajustada.xlsx', index=False, header=False)