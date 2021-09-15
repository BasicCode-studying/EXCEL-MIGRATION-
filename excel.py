import pandas as pd
from time import sleep
import os


# contatos_df = pd.read_excel("C:/Users/user/Desktop/Auto/implementação (2).xlsm", index_col=None, na_values=["NA"], usecols="B, H")
# contatos_df = contatos_df.loc[1]
contatos_df = pd.read_excel("C:/Users/user/Desktop/Auto/implementação (2).xlsm")


for i, mensagem in enumerate(contatos_df['Número do processo:']):
    filesequence = + 1
    Honrarios = contatos_df.loc[i, 'Honorários %']
    Numero_Processo = contatos_df.loc[i, "Número do processo:"]
    Numero_Incidente = contatos_df.loc[i, "Numero do Incidente:"]
    Termo_Inicial = contatos_df.loc[i, "TERMO INICIAL"]
    Termo_Final = contatos_df.loc[i, "TERMO FINAL"]
    Foro = contatos_df.loc[i, "Foro:"]
    Score = contatos_df.loc[i, "Score:"]
    Data_Base = contatos_df.loc[i, "Data base:"]
    Data_Decisao = contatos_df.loc[i, "DataDecisão:"]
    Liquido = contatos_df.loc[i, "Principal líquido:"]
    Juros = contatos_df.loc[i, "Juros moratórios:"]
    Valor_negociavel = contatos_df.loc[i, "Valor Negociável (-30%) "]
    Entidade = contatos_df.loc[i, "Entidade devedora:"]
    Requerente = contatos_df.loc[i, "Requerente :"]
    CPF = contatos_df.loc[i, 'CPF:']
    print(f"PROCESSOS POR CREDORES: - {Honrarios} - {Numero_Processo} - {Numero_Incidente} - "
          f"{Termo_Inicial} - {Termo_Final} - {Foro} - "
          f"{Score} - {Data_Base} - {Data_Decisao} - "
          f"{Liquido} - {Juros} - {Valor_negociavel} - "
          f"{Entidade} - {Requerente} - {CPF}")
    migrando = pd.DataFrame({"Honorários": [Honrarios],
                       "NUMERO_PROCESSO": [Numero_Processo],
                       "NUMERO_INCIDENTE": [Numero_Incidente],
                       "TERMO_INICIAL": [Termo_Inicial],
                       "TERMO_FINAL": [Termo_Final],
                       "FORO": [Foro],
                       "SCORE": [Score],
                       "DATA_BASE": [Data_Base],
                       "DATA_DECISÃO": [Data_Decisao],
                       "Principal_Liquido": [Liquido],
                       "JUROS_MORATÓRIO": [Juros],
                       "VALOR_NEGOCIÁVEL": [Valor_negociavel],
                       "ENTIDADE_DEVEDORA": [Entidade],
                       "REQUERENTE": [Requerente],
                       "CPF": [CPF]})

    migrando.to_excel(f"MIGRANDO{filesequence}.xlsx")
    filesequence += 1
    sleep(2)
