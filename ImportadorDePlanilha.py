import pandas as pd
import numpy as np


df = pd.read_excel("orcamento.xlsx", skiprows=4)
tamanho_plan = len(df.index)
lista = []

for i in range(0,tamanho_plan,1): # identificando ultima linha com item e apagando linhas vazias (célula vazia)
    linha = df.loc[i,'Item']
    cond = (pd.isnull(linha))
    if cond == True:
       linha = i
       df = df.drop(linha)
       
for i in range(0,tamanho_plan,1): # formatando coluna da itemização

    l1 = df.loc[i,'Item']
    l1 = str(l1)
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

    elif len(array) == 4:
        item0 = array[0]
        itemformat0 = int(item0)
        istr=str(itemformat0)

        item1 = array[1]
        itemformat1 = int(item1)
        istr1=str(itemformat1)

        item2 = array[2]
        itemformat2 = int(item2)
        istr2=str(itemformat2)

        item3 = array[3]
        itemformat3 = int(item3)
        istr3=str(itemformat3)

        ifin2=(istr+'.'+istr1+'.'+istr2+'.'+istr3)
        lista.append(ifin2)
        
for i in range(0,150,1): # alterando itens da coluna inteira
    df.loc[i,'Item'] = lista[i]

print(df)
df.to_excel('Planilha Ajustada.xlsx')
