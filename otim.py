import pandas as pd
import numpy as np
import FormatBancos

df = pd.read_excel("orcamento.xlsx", header=None)
tamanho_plan = len(df.index)


lerr = int(input("Linha primeiro item: "))
l0 = lerr-1
lf = int(input("Linha ultimo item: "))
lista = []

for i in range(l0,lf):
    linha = df.loc[i,0]

for i in range(l0,lf): # formatando coluna da itemização

    l1 = df.loc[i,0]
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

        ifin=(istr+'.'+istr1+'.'+istr2)
        lista.append(ifin)

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

        ifin=(istr+'.'+istr1+'.'+istr2+'.'+istr3)
        lista.append(ifin)

    elif len(array) == 5:
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

        item4 = array[4]
        itemformat4 = int(item4)
        istr4=str(itemformat4)

        ifin=(istr+'.'+istr1+'.'+istr2+'.'+istr3+'.'+istr4)
        lista.append(ifin)

for i in range(l0,lf): # alterando itens da coluna inteira
    df.loc[i,0] = lista[i]