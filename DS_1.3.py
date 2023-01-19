#-------------------------------------------------------------------------------
# Name:        DS 1.3 - Projeto Carros - POLO
# Purpose:
#
# Author:      Ivo
#
# Created:      23dez2022
# Copyright:   (c) Ivo 2022
# Licence:     <your licence>
#-------------------------------------------------------------------------------

import requests
import camelot
import pandas as pd
import os
import numpy as np
import plotly.express as px

print("LOSS !")

path_pdf = "D:\\Docs\\Python\\Projeto Carros\\pdf_original_database\\"
path_tables = "D:\\Docs\\Python\\Projeto Carros\\tables\\"
path_graphs = "D:\\Docs\\Python\\Projeto Carros\\graphics\\"

ano = 2022
mes = 9
ano_mes = str(ano) + "_" + f'{mes:02d}'
print(ano_mes)


def read_pdf(pg, table):
    """ Lê o arquivo pdf e gera um Dataframe

        Parâmetros
        pg : numeral em string, número da página do arquivo pdf
        table : int, número da tabela da página a ser lida

        Returns
        Dataframe do mês
        """

    table_read = camelot.read_pdf(path_pdf + ano_mes + "_2.pdf", pages = pg, flavor='stream')
    df_read = table_read[table].df
    return df_read


def clean_df (df):
    """Prepara e formata o Dataframe para ter seus dados processados.

        Parâmetro
        df : o Dataframe gerado pela função read_pdf

        Returns
        Dataframe do mês
    """

    df_cleaned = df.replace("", np.nan)
    df_cleaned.dropna(axis=0, inplace=True)
    df_cleaned = df_cleaned.set_axis(["Modelo", str(ano_mes)], axis=1)     # define nome das colunas.
    df_cleaned = df_cleaned[df_cleaned["Modelo"].str.contains("Modelo") == False]    # remove linha com Modelo
    df_cleaned.reset_index(inplace=True, drop=True)
    df_cleaned = df_cleaned.astype("string")
    df_cleaned[str(ano_mes)] = df_cleaned[str(ano_mes)].str.replace('.', '', regex=False)      # remover o ponto de milhar na coluna ano_mes
    df_cleaned[str(ano_mes)] = df_cleaned[str(ano_mes)].astype("int64")
    return df_cleaned


def save_df (df, xlsx_name):
    """Salva o Dataframe em xlsx e o abre para inspeção visual.

        Parâmetros
        df : Dataframe a salvar.
        xlsx_name : string, nome do arquivo xlsx a salvar.
    """

    xls = (path_tables + ano_mes + "_" + xlsx_name + ".xlsx")
    df.to_excel(xls)
    os.startfile(xls)


def master (master_name, df):
    """ Lê o master do mês anterior, acrescenta o novo mês, prepara o df
        e o salva em xlsx.

        Parâmetros
        master_name : o df master a ser aberto.
        df : o df do mês a acrescentar.

        Returns
        O Master Dataframe
    """

    master = pd.read_excel(path_tables + ano_mes_anterior + "_" + master_name + ".xlsx")
    master = pd.merge(master, df, how="outer", on="Modelo")      # mesclar com o novo mês
    master = master.fillna(0)                                         # substituir os NaN por zero
    master.loc[:, "2017_11":str(ano_mes)] = master.loc[:, "2017_11":str(ano_mes)].astype("int")         # transforma todos dados em inteiro.
    master.set_index("Modelo", inplace = True)               # deixar modelo como index das linhas.
    master = master.loc[(master!=0).any(axis=1)]       # remove linhas com zero qnt
    master.to_excel(path_tables + ano_mes + "_" + master_name + ".xlsx")
    return master


def master_24m (df):
    """ Deixa o df master com apenas os últimos 24 meses.

        Parâmetros
        df : o df master a ser reduzido.

        Returns
        O Dataframe reduzido
    """

    df_24m = df.loc[:, ano_mes_24m : ano_mes]
    df_24m = df_24m.loc[(df_24m!=0).any(axis=1)]       # remove linhas com zero qnt
    return df_24m


def ranking (df, nome):
    """ Gera arquivo xlsx ordenado pelos modelos mais vendido.

        Parâmetros
        df : o df master para gerar o rank.
        nome : nome do arquivo xlsx a salvar.
        """

    df["total_modelo"] = df.sum(axis=1)
    rank = df[["total_modelo"]].sort_values(by=["total_modelo"], ascending=False).reset_index()
    rank.to_excel(path_tables + ano_mes + "_" + nome + ".xlsx")
    df.drop(labels = ["total_modelo"], axis = 1, inplace = True)


def transpor (df):
    """Transpoem o df, gera novas colunas e o prepara gerar os gráficos.

        Parâmetros
        df : o df base dos gráficos.

        Returns
        O df preparado
        """

    df_t = df.transpose()
    df_t['TOTAIS'] = df_t.sum(axis=1)
    df_t["Média móvel TOTAIS"] = df_t['TOTAIS'].rolling(window=6, min_periods=1).mean().astype(int)
    df_t["Média móvel Polo"] = df_t["VW/POLO"].rolling(window=6, min_periods=1).mean()
    df_t["% Polo/Total"] = df_t["VW/POLO"] * 100 / df_t["TOTAIS"]
    df_t["média móvel %"]= df_t["% Polo/Total"].rolling(window=6, min_periods=1).mean()
    return df_t


def graphs (df_col, gr_title, name, label):
    """ Gera os gráfico

        Parâmetros
        df_col : lista das colunas do df de cada linha do gráfico.
        gr_title : string com o título do gráfico
        name : nome do gráfico a salvar
        label : dicionário o nome dos eixos do gráfico.
        """

    graph = px.line(df_col, title=gr_title,
    color_discrete_sequence=["red", "green", "blue"],
    labels=label, markers=True)
    graph.update_layout(autosize=False, width=1200, height=600,
    margin=dict(l=30, r=20, t=40, b=20),
    legend=dict(yanchor="top", y=0.99, xanchor="left", x=0.01),
    font_family="IBM Plex Mono Medium")
    graph.write_image(path_graphs + ano_mes + "_" + name + ".png")


#       mes anterior
if mes != 1:
    ano_mes_anterior = str(ano) + "_" + f'{(mes - 1):02d}'
else:
    ano_mes_anterior = str(ano - 1) + "_" + "12"

#       24 meses antes
if mes != 12:
    ano_mes_24m = str(ano - 2) + "_" + f'{(mes + 1):02d}'
else:
    ano_mes_24m = str(ano - 1) + "_" + "1"

file_url = ("http://www.fenabrave.org.br//portal//files//" + ano_mes + "_2.pdf")
local_file = (path_pdf + ano_mes + "_2.pdf")
response = requests.get(file_url, headers={"User-Agent": "Mozilla/5.0"})
open(local_file, "wb").write(response.content)
print(response)
os.startfile(local_file)

df_cars_read = read_pdf("6", 0)
df_cars_read = df_cars_read.iloc[:, 1:3]
#print(df_cars_read.head())

df_seg_read = read_pdf("11", 1)
df_seg_read = df_seg_read.iloc[4:16, [1, 3]]
#print(df_seg_read)

df_cars_cleaned = clean_df(df_cars_read)
df_seg_cleaned = clean_df(df_seg_read)

save_df(df_cars_cleaned, "df_cars")
save_df(df_seg_cleaned, "df_seg")

df_master = master("df_master", df_cars_cleaned)
df_seg_master = master("df_seg_master", df_seg_cleaned)
#print(df_master.head())
#print(df_seg_master.head())

df_master_24m = master_24m(df_master)
df_seg_master_24m = master_24m(df_seg_master)

rank_all_life = ranking(df_master, "rank_all_life")
rank_24m = ranking(df_master, "rank_24m")
rank_seg_all_life = ranking(df_seg_master, "rank_seg_all_life")
rank_seg_24m = ranking(df_master, "rank_seg_24m")

df_master_t = transpor(df_master)
df_master_24m_t = transpor(df_master_24m)
df_seg_master_t = transpor(df_seg_master)
df_seg_master_24m_t = transpor(df_seg_master_24m)

labels_qnt ={"value": "Unidades x 1.000",  "index": "Ano_mês"}
labels_per_cent = {"value": "%",  "index": "Ano_mês"}
gr01 = graphs(df_master_t[["VW/POLO", "Média móvel TOTAIS", "TOTAIS"]],
"Carros Novos Emplacados", "gr01 - Totais", labels_qnt)
gr02 = graphs(df_master_24m_t[["VW/POLO", "Média móvel TOTAIS", "TOTAIS"]],
"Carros Novos Emplacados Últimos 24 Meses", "gr02 - Totais 24m", labels_qnt)
gr03 = graphs(df_master_t[["VW/POLO", "Média móvel Polo"]],
"VW/POLO Novos Emplacados", "gr03 - Polo", labels_qnt)
gr04 = graphs(df_master_24m_t[["VW/POLO", "Média móvel Polo"]],
"VW/POLO Novos Emplacados Últimos 24 meses", "gr04 - Polo 24m", labels_qnt)
gr05 = graphs(df_master_t[["% Polo/Total", "média móvel %"]],
"VW/POLO % do Total de Novos Emplacados", "gr05 - % Polo", labels_per_cent)
gr06 = graphs(df_master_24m_t[["% Polo/Total", "média móvel %"]],
"VW/POLO % do Total de Novos Emplacados Últimos 24 meses", "gr06 - % Polo 24m", labels_per_cent)
gr07 = graphs(df_seg_master_t[["VW/POLO", "Média móvel TOTAIS", "TOTAIS"]],
"Hatchs Pequens Novos Emplacados", "gr07 - Totais seg", labels_qnt)
gr08 = graphs(df_seg_master_24m_t[["VW/POLO", "Média móvel TOTAIS", "TOTAIS"]],
"Hatchs Pequens Novos Emplacados Últimos 24 meses", "gr08 - Totais seg 24m", labels_qnt)
gr09 = graphs(df_seg_master_t[["% Polo/Total", "média móvel %"]],
"VW/POLO % do Hatchs Pequenos de Novos Emplacados", "gr09 - % Polo seg", labels_per_cent)
gr10 = graphs(df_seg_master_24m_t[["% Polo/Total", "média móvel %"]],
"% de VW/POLO dos Hatchs Pequenos Novos Emplacados Últimos 24 meses", "gr10 - % Polo seg 24m", labels_per_cent)

    # FIM

print("ES HAT GETAN")

