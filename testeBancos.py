import pandas as pd
import numpy as np

df = pd.read_excel("orcamento.xlsx", skiprows=4)
tamanho_plan = len(df.index)

for i in range(0,tamanho_plan,1):
    bc = df.loc[i,'Banco']
    
    if bc == 'SINAPI-I' or bc == 'sinapi-i' or bc == 'sinapi-c' or bc == 'SINAPI-C' or bc =='sinapi':
        df.loc[i,'Banco'] = "SINAPI"
    elif bc == 'SBC' or bc == 'sbc':
        df.loc[i,'Banco'] = 'SBC'

        
print(df['Banco']) 