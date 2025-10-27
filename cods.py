import pandas as pd
import numpy as np

df = pd.read_excel("orcamento.xlsx", skiprows=4)
tamanho_plan = len(df.index)


for i in range(0,tamanho_plan,1):
    bc = df.loc[i,'Banco']
    if bc == "SINAPI":
        code = (df.loc[i,"Código"])
        codel = code.strip()
        df.loc[i,'Código'] = codel
        if len(codel) == 5 or len(codel) == 6:
            print("formato correto")
        else:
            print("erro")
            df.loc[i,"Código"] = "Erro"
print(df)

df.to_excel('Planilha Ajustada.xlsx', index=False)