import numpy as np
import pandas as pd

itens = []
lista_int = [int(i) for i in itens]
df = pd.read_excel('orcamento.xlsx', skiprows=4)


for i in range(0,15,1):
  i = df.loc[i,"Item"]
  i_sep = i.split(".")
  lista_int.append(i_sep)
  
print(lista_int)
