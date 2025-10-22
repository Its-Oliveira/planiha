import pandas as pd
import numpy as np
import openpyxl

df = pd.read_excel("orcamento.xlsx", skiprows=4)

print(df['Item'])
lista = []
for i in range(0,12,1): # formatando coluna da itemização
    l1 = df.loc[i,'Item']
    i_sep = l1.split('.')
    array = np.array(i_sep)
    if len(array) == 1:
        item0 = array[0]
        itemformat0 = int(item0)
        istr=str(itemformat0)
        lista.append(istr)

    elif len(array) == 2:
        item0 = array[0]
        itemformat0 = int(item0)
        istr=str(itemformat0)
        item1 = array[1]
        itemformat1 = int(item1)
        istr1=str(itemformat1)
        ifin=(istr+'.'+istr1)
        lista.append(ifin)

    elif len(array) == 3:
        item0 = array[0]
        itemformat0 = int(item0)
        istr=str(itemformat0)

        item1 = array[1]
        itemformat1 = int(item1)
        istr1=str(itemformat1)

        item2 = array[2]
        itemformat2 = int(item2)
        istr2=str(itemformat2)
        ifin2=(istr+'.'+istr1+'.'+istr2)
        lista.append(ifin2)
        

for i in range(0,12,1): # alterando itens da coluna inteira
    df.loc[i,'Item'] = lista[i]
print(df['Item'])

df.to_excel('teste.xlsx')