import pandas as pd

df = pd.read_excel("orcamento.xlsx", skiprows=4)
tamanho_plan = len(df.index)


def SyntaxBancos(df):
    for i in range(0,150,1):
        bc = df.loc[i,'Banco']
        
        if bc == 'SINAPI-I' or bc == 'sinapi-i' or bc == 'sinapi-c' or bc == 'SINAPI-C' or bc =='sinapi' or bc=='SINAPI':
            df.loc[i,'Banco'] = "SINAPI"

        elif bc == 'SBC' or bc == 'sbc':
            df.loc[i,'Banco'] = 'SBC'

        elif bc=="CPOS" or bc=='cpos' or bc=='CDHU' or bc =='cdhu' or bc == 'CPOS/CDHU' or bc =='cpos/cdhu':
            df.loc[i,'Banco'] = 'CPOS'

        elif bc =="SICRO3" or bc == 'sicro3' or bc == 'SICRO' or bc == 'sicro':
            df.loc[i,'Banco'] = 'SICRO3' 

        elif bc == 'ORSE' or bc=='orse':
            df.loc[i,'Banco'] = 'ORSE'

        elif bc == 'sedop' or bc=='SEDOP':
            df.loc[i,'Banco'] = 'SEDOP'
            
        elif bc =='SEINFRA' or bc=='seinfra':
            df.loc[i,'Banco'] = 'SEINFRA'

        elif bc =='setop' or bc=='SETOP':
            df.loc[i,'Banco'] = 'SETOP'

        elif bc =='IOPES' or bc=='iopes':
            df.loc[i,'Banco'] = 'IOPES'

        elif bc =='SIURB' or bc=='siurb':
            df.loc[i,'Banco'] = 'SIURB'

        elif bc =='SIURB INFRA' or bc=='siurb infra' or bc=='SIURB infra' or bc=='siurb INFRA':
            df.loc[i,'Banco'] = 'SIURB INFRA'

        elif bc =='SUDECAP' or bc=='sudecap':
            df.loc[i,'Banco'] = 'SUDECAP'

        elif bc =='FDE' or bc=='fde':
            df.loc[i,'Banco'] = 'FDE'

        elif bc =='AGESUL' or bc=='agesul':
            df.loc[i,'Banco'] = 'AGESUL'

        elif bc =='FDE' or bc=='fde':
            df.loc[i,'Banco'] = 'FDE'
        
        elif bc =='AGETOP CIVIL' or bc=='agetop civil' or bc=='AGETOP civil' or bc=='agetop CIVIL':
            df.loc[i,'Banco'] = 'AGETOP CIVIL'

        elif bc =='AGETOP RODOVIARIA' or bc=='agetop rodoviaria' or bc=='AGETOP rodoviaria' or bc=='agetop RODOVIARIA':
            df.loc[i,'Banco'] = 'AGETOP RODOVIARIA'

        elif bc =='CAEMA' or bc=='caema':
            df.loc[i,'Banco'] = 'CAEMA'

        elif bc =='EMBASA' or bc=='embasa':
            df.loc[i,'Banco'] = 'EMBASA'

        elif bc =='CAERN' or bc=='caern':
            df.loc[i,'Banco'] = 'CAEMA'

        elif bc =='COMPESA' or bc=='compesa':
            df.loc[i,'Banco'] = 'COMPESA'

        elif bc =='EMOP' or bc=='emop':
            df.loc[i,'Banco'] = 'EMOP'

    return df

SyntaxBancos(df)
df.to_excel('Planilha Ajustada.xlsx')

   

        