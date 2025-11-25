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
            if len(codel) == 5 or len(codel) == 6 or len(codel) == 8:
                df.loc[i,1] = codel
            else:
                df.loc[i,1] = "Erro, formato não reconhecido pela SINAPI"

        elif bc == "SBC":
            code = str(df.loc[i,1])
            codel = (code.strip())
            df.loc[i,1] = codel
            if len(codel) == 6:
                df.loc[i,1] = codel
            else:
                df.loc[i,1] = "Erro, formato não reconhecido pela SBC"


        elif bc == "CPOS":
            code = str(df.loc[i,1])
            codel = (code.strip())
            df.loc[i,1] = codel
            if len(codel) == 9 or len(codel) == 15:
                df.loc[i,1] = codel
            else:
                df.loc[i,1] = "Erro, formato não reconhecido pela CPOS"

        elif bc == "SICRO":
            code = str(df.loc[i,1])
            codel = (code.strip())
            df.loc[i,1] = codel
            if len(codel) == 7 or len(codel) == 5:
                df.loc[i,1] = codel
            else:
                df.loc[i,1] = "Erro, formato não reconhecido pela SICRO"

        elif bc == "SETOP":
            code = str(df.loc[i,1])
            codel = (code.strip())
            df.loc[i,1] = codel
            if len(codel) == 8 or len(codel) == 11:
                df.loc[i,1] = codel
            else:
                df.loc[i,1] = "Erro, formato não reconhecido pela SETOP"

        elif bc == "IOPES":
            code = str(df.loc[i,1])
            codel = (code.strip())
            df.loc[i,1] = codel
            if len(codel) == 6:
                df.loc[i,1] = codel
            else:
                df.loc[i,1] = "Erro, formato não reconhecido pela IOPES"

        elif bc == "SIURB":
            code = str(df.loc[i,1])
            codel = (code.strip())
            df.loc[i,1] = codel
            if len(codel) == 7 or len(codel) == 5:
                df.loc[i,1] = codel
            else:
                df.loc[i,1] = "Erro, formato não reconhecido pela SIURB"

        elif bc == "SIRUB INFRA":
            code = str(df.loc[i,1])
            codel = (code.strip())
            df.loc[i,1] = codel
            if len(codel) == 7 or len(codel) == 5:
                df.loc[i,1] = codel
            else:
                df.loc[i,1] = "Erro, formato não reconhecido pela SIURB INFRA"

        elif bc == "SUDECAP":
            code = str(df.loc[i,1])
            codel = (code.strip())
            df.loc[i,1] = codel
            if len(codel) == 8:
                df.loc[i,1] = codel
            else:
                df.loc[i,1] = "Erro, formato não reconhecido pela SUDECAP"

        elif bc == "FDE":
            code = str(df.loc[i,1])
            codel = (code.strip())
            df.loc[i,1] = codel
            if len(codel) == 7 or len(codel) == 9:
                df.loc[i,1] = codel
            else:
                df.loc[i,1] = "Erro, formato não reconhecido pela FDE"

        elif bc == "EMOP":
            code = str(df.loc[i,1])
            codel = (code.strip())
            df.loc[i,1] = codel
            if len(codel) == 13 or len(codel) == 5:
                df.loc[i,1] = codel
            else:
                df.loc[i,1] = "Erro, formato não reconhecido pela EMOP"

        elif bc == "SCO":
            code = str(df.loc[i,1])
            codel = (code.strip())
            df.loc[i,1] = codel
            if len(codel) == 13 or len(codel) == 9:
                df.loc[i,1] = codel
            else:
                df.loc[i,1] = "Erro, formato não reconhecido pela SCO"
                print(len(codel))

        elif bc == "ORSE":
            code = str(df.loc[i,1])
            codel = (code.strip())
            df.loc[i,1] = codel
            if len(codel) == 4 or len(codel) == 5 or len(codel) == 15 or len(codel) == 3:
                df.loc[i,1] = codel
            else:
                df.loc[i,1] = "Erro, formato não reconhecido pela ORSE"

        elif bc == "SEINFRA":
            code = str(df.loc[i,1])
            codel = (code.strip())
            df.loc[i,1] = codel
            if len(codel) == 5:
                df.loc[i,1] = codel
            else:
                df.loc[i,1] = "Erro, formato não reconhecido pela SEINFRA"

        elif bc == "CAEMA":
            code = str(df.loc[i,1])
            codel = (code.strip())
            df.loc[i,1] = codel
            if len(codel) == 6 or len(codel) == 10:
                df.loc[i,1] = codel
            else:
                df.loc[i,1] = "Erro, formato não reconhecido pela CAEMA"

        elif bc == "EMBASA":
            code = str(df.loc[i,1])
            codel = (code.strip())
            df.loc[i,1] = codel
            if len(codel) == 8 or len(codel) == 10:
                df.loc[i,1] = codel
            else:
                df.loc[i,1] = "Erro, formato não reconhecido pela EMBASA"

        elif bc == "CAERN":
            code = str(df.loc[i,1])
            codel = (code.strip())
            df.loc[i,1] = codel
            if len(codel) == 3 or len(codel) == 4 or len(codel) == 5 or len(codel) == 6 or len(codel) == 7:
                df.loc[i,1] = codel
            else:
                df.loc[i,1] = "Erro, formato não reconhecido pela CAERN"

        elif bc == "COMPESA":
            code = str(df.loc[i,1])
            codel = (code.strip())
            df.loc[i,1] = codel
            if len(codel) == 9 or len(codel) == 13 or len(codel) == 15:
                df.loc[i,1] = codel
            else:
                df.loc[i,1] = "Erro, formato não reconhecido pela COMPESA"

        elif bc == "AGESUL":
            code = str(df.loc[i,1])
            codel = (code.strip())
            df.loc[i,1] = codel
            if len(codel) == 4 or len(codel) == 10:
                df.loc[i,1] = codel
            else:
                df.loc[i,1] = "Erro, formato não reconhecido pela AGESUL"

        elif bc == "AGETOP CIVIL":
            code = str(df.loc[i,1])
            codel = (code.strip())
            df.loc[i,1] = codel
            if len(codel) == 6 or len(codel) == 4:
                df.loc[i,1] = codel
            else:
                df.loc[i,1] = "Erro, formato não reconhecido pela AGETOP CIVIL"

        elif bc == "AGETOP RODOVIARIA":
            code = str(df.loc[i,1])
            codel = (code.strip())
            df.loc[i,1] = codel
            if len(codel) == 5:
                df.loc[i,1] = codel
            else:
                df.loc[i,1] = "Erro, formato não reconhecido pela AGETOP RODOVIARIA"

        elif bc == "SEDOP":
            code = str(df.loc[i,1])
            codel = (code.strip())
            df.loc[i,1] = codel
            if len(codel) == 8 or len(codel) == 6:
                df.loc[i,1] = codel
            else:
                df.loc[i,1] = "Erro, formato não reconhecido pela SEDOP"

        elif bc == "DERPR":
            code = str(df.loc[i,1])
            codel = (code.strip())
            df.loc[i,1] = codel
            if len(codel) == 6:
                df.loc[i,1] = codel
            else:
                df.loc[i,1] = "Erro, formato não reconhecido pela DERPR"

        

codes(df,l0,lf)
        




df.to_excel('Planilha Ajustada.xlsx', index=False, header=False)