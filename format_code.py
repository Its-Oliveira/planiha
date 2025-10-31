import pandas as pd
import numpy as np

df = pd.read_excel("orcamento.xlsx", header=None)
tamanho_plan = len(df.index)


def codes(l0,lf):
    for i in range(l0,lf):
        bc = df.loc[i,2]
        if bc == "SINAPI":
            code = (df.loc[i,"Código"])
            codel = code.strip()
            df.loc[i,'Código'] = codel
            if len(codel) == 5 or len(codel) == 6:
                print("formato correto")
                print(codel)
            else:
                print("erro")
                df.loc[i,"Código"] = "Erro"
print(df)

df.to_excel('Planilha Ajustada.xlsx', index=False, header=False)