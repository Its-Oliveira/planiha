import pandas as pd


df = pd.read_excel("orcamento.xlsx", header=None)
tamanho_plan = len(df.index)


def SyntaxBancos(df, l0, lf):
    for i in range(l0,lf):
        bc = df.loc[i,2]
        
        if bc == 'SINAPI-I' or bc == 'sinapi-i' or bc == 'sinapi-c' or bc == 'SINAPI-C' or bc =='sinapi' or bc=='SINAPI':
            df.loc[i,2] = "SINAPI"

        elif bc == 'SBC' or bc == 'sbc':
            df.loc[i,2] = 'SBC'

        elif bc=="CPOS" or bc=='cpos' or bc=='CDHU' or bc =='cdhu' or bc == 'CPOS/CDHU' or bc =='cpos/cdhu':
            df.loc[i,2] = 'CPOS'

        elif bc =="SICRO3" or bc == 'sicro3' or bc == 'SICRO' or bc == 'sicro':
            df.loc[i,2] = 'SICRO3' 

        elif bc == 'ORSE' or bc=='orse':
            df.loc[i,2] = 'ORSE'

        elif bc == 'sedop' or bc=='SEDOP':
            df.loc[i,2] = 'SEDOP'
            
        elif bc =='SEINFRA' or bc=='seinfra':
            df.loc[i,2] = 'SEINFRA'

        elif bc =='setop' or bc=='SETOP':
            df.loc[i,2] = 'SETOP'

        elif bc =='IOPES' or bc=='iopes':
            df.loc[i,2] = 'IOPES'

        elif bc =='SIURB' or bc=='siurb':
            df.loc[i,2] = 'SIURB'

        elif bc =='SIURB INFRA' or bc=='siurb infra' or bc=='SIURB infra' or bc=='siurb INFRA':
            df.loc[i,2] = 'SIURB INFRA'

        elif bc =='SUDECAP' or bc=='sudecap':
            df.loc[i,2] = 'SUDECAP'

        elif bc =='FDE' or bc=='fde':
            df.loc[i,2] = 'FDE'

        elif bc =='AGESUL' or bc=='agesul':
            df.loc[i,2] = 'AGESUL'

        elif bc =='FDE' or bc=='fde':
            df.loc[i,2] = 'FDE'
        
        elif bc =='AGETOP CIVIL' or bc=='agetop civil' or bc=='AGETOP civil' or bc=='agetop CIVIL':
            df.loc[i,2] = 'AGETOP CIVIL'

        elif bc =='AGETOP RODOVIARIA' or bc=='agetop rodoviaria' or bc=='AGETOP rodoviaria' or bc=='agetop RODOVIARIA':
            df.loc[i,2] = 'AGETOP RODOVIARIA'

        elif bc =='CAEMA' or bc=='caema':
            df.loc[i,2] = 'CAEMA'

        elif bc =='EMBASA' or bc=='embasa':
            df.loc[i,2] = 'EMBASA'

        elif bc =='CAERN' or bc=='caern':
            df.loc[i,2] = 'CAEMA'

        elif bc =='COMPESA' or bc=='compesa':
            df.loc[i,2] = 'COMPESA'

        elif bc =='EMOP' or bc=='emop':
            df.loc[i,2] = 'EMOP'

    return df


def format_itemizacao(valor):
    try:
        partes = str(valor).split(".")
        partes_formatadas = [str(int(p)) for p in partes if p.strip() != ""]
        return ".".join(partes_formatadas)
    
    except ValueError:
        # Caso alguma parte não seja número, retorna o valor original
        return valor
    
    
def codes(df,l0,lf):
    for i in range(l0,lf):
        bc = df.loc[i,2]
        if bc == "SINAPI":
            code = str(df.loc[i,1])
            codel = (code.strip())
            df.loc[i,1] = codel
            if len(codel) == 5 or len(codel) == 6 or len(codel) == 8:
                df.loc[i,1] = codel
            else:
                df.loc[i,1] = "Erro, formato não reconhecido pela SINAPI"

df.to_excel('Planilha Ajustada.xlsx')

   

        