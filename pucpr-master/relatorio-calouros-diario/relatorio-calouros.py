'''
Script recebe um arquivo Excel com relação dos alunos e gera um relatório sobre os calouros, 
realizando uma contagem sobre os tipos de ingressso  de cada aluno.

Desenvolvido por Guilherme Hann e Vinicius Tozo
Última atualização: 20/07/2020
'''


# coding:utf8
import pandas as pd
import datetime
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import pytz

dicionario = {}

Tk().withdraw()
relatorio_calouros = askopenfilename(
    filetypes=[('Arquivo excel', '.xlsx')],
    title='Selecione a relação de alunos')

df_alunos = pd.read_excel(relatorio_calouros, index=False, skiprows=1)

df_alunos = df_alunos[df_alunos.Série.eq("1º Período")]

df_alunos["Centro de Resultado"] = df_alunos["Centro de Resultado"].fillna(0)

for index, line in df_alunos.iterrows():
    key = line["Estabelecimento"] + ";" + str(int(line["Centro de Resultado"])) \
          + ";" + line["Curso"] + ";" + line["Série"]
    if key in dicionario:
        dicionario[key][0] += 1
    else:
        dicionario[key] = [1, 0, 0, 0, 0, 0]

    if line["Tipo de Ingresso"] == "Enem":
        dicionario[key][1] += 1
    elif line["Tipo de Ingresso"] == "Portador de Diploma":
        dicionario[key][2] += 1
    elif line["Tipo de Ingresso"] == "Processo Seletivo":
        dicionario[key][3] += 1
    elif line["Tipo de Ingresso"] == "Transferência":
        dicionario[key][4] += 1
    else:
        dicionario[key][5] += 1

lista_dicionario = []

for linha in dicionario:
    lista_dicionario.append(
        [linha.split(";")[0], linha.split(";")[1], linha.split(";")[2], linha.split(";")[3], dicionario[linha][0],
         dicionario[linha][1], dicionario[linha][2], dicionario[linha][3], dicionario[linha][4], dicionario[linha][5]])

df = pd.DataFrame(lista_dicionario,
                  columns=["Estabelecimento", "Centro de Resultado", "Curso", "Série", "Total de estudantes",
                           "Enem", "Portador de Diploma", "Processo Seletivo", "Transferência",
                           "Tipo de ingresso vazio"])

utc_now = pytz.utc.localize(datetime.datetime.utcnow())
dt_now = utc_now.astimezone(pytz.timezone("America/Sao_Paulo"))

nome_arquivo = 'arquivo_calouros ' + str(dt_now.strftime("%Y-%m-%d %H-%M")) + '.xlsx'

writer = pd.ExcelWriter(nome_arquivo, engine='xlsxwriter')
df.to_excel(writer, sheet_name='Sheet1', index=False)
writer.save()
