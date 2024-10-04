'''
Script unifica todas as informações dos alunos e as disciplinas que ele cursou em cada ano

Desenvolvido por Guilherme Hann e Vinicius Tozo
Última atualização: 17/07/2020
'''

import pandas as pd
import xlrd
import datetime
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import xlsxwriter
import pytz


def createDf(endereco):
    global df
    df = pd.read_excel(endereco, index=False)
    df = deleteTrashRows(df)
    return df


def deleteTrashRows(df):
    del df['Codigo PA Matriz']
    del df['Modulacao Teorica']
    del df['Modulacao Pratica']
    del df['Codigo PA Oferta']
    del df['Situacao Academica']
    del df['Divisão Turma Teórica']
    del df['Divisão Turma Prática']
    del df['Divisão Turma Tutoria']
    del df['Total HA']
    del df['ano']
    del df['semestre']
    del df['sigla']
    del df['Nome do Turno']
    del df['Tipo da Turma']
    del df['Data Nascimento']
    del df['Tipo dispensa']
    del df['Tipo Entrada']
    del df['situacao_contratual']
    del df['Ingresso aluno']
    del df['Media final']
    del df['Faltas']
    del df['(%) Presenca']
    del df['Exame Pendente']
    del df['Nota 1']
    del df['Nota 2']
    del df['Nota 3']
    del df['Nota 4']
    del df['Média I']
    del df['Nota 5']
    del df['Nota 6']
    del df['Nota 7']
    del df['Nota 8']
    del df['Média II']
    del df['Nota Avaliacao Final']
    del df['Data Original']
    del df['cpfcnpj']
    del df['Cod. Curso Oferta']
    del df['Nome do Curso Oferta']
    return df


df1 = createDf('../in/certificados_intermediarios_ed_cod/Relatorio 2019-1 versao 2.xlsx')
df2 = createDf('../in/certificados_intermediarios_ed_cod/relatorio_Relatório de alunos e disciplinas - 2019-1.xlsx')
df3 = createDf('../in/certificados_intermediarios_ed_cod/relatorio_Relatório de alunos e disciplinas  2018-1.xlsx')
df4 = createDf('../in/certificados_intermediarios_ed_cod/relatorio_Relatório de alunos e disciplinas  2018-2.xlsx')
df5 = createDf('../in/certificados_intermediarios_ed_cod/relatorio_Relatório de alunos e disciplinas  2019-2.xlsx')
df6 = createDf('../in/certificados_intermediarios_ed_cod/relatório_Relatorio de alunos e disciplinas 2020-1.xlsx')
df7 = createDf('../in/certificados_intermediarios_ed_cod/Relatorio 2019-1 versao 4.xlsx')

frames = [df1, df2, df3, df4, df5, df6, df7]
resultado = pd.concat(frames)

resultado.drop(resultado.loc[resultado['Status Aprovacao'] == 'Não processado'].index, inplace=True)
resultado.drop(resultado.loc[resultado['Status Aprovacao'] == 'Horas não cumpridas'].index, inplace=True)
resultado.drop(resultado.loc[resultado['Status Aprovacao'] == 'Reprovado por frequência'].index, inplace=True)
resultado.drop(resultado.loc[resultado['Status Aprovacao'] == 'Reprovado por nota'].index, inplace=True)
resultado.drop(resultado.loc[resultado['Status Aprovacao'] == 'Reprovado por nota e frequência'].index, inplace=True)

# cols = ['Billing Address Street 1', 'Billing Address Street 2','Billing Company']
# for col in cols:
#     if col in df.columns:
#         df = df.drop(columns=col, axis=1)

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('arquivo_exportado.xlsx', engine='xlsxwriter')

# Convert the dataframe to an XlsxWriter Excel object.# Dados Essenciais
resultado.to_excel(writer, sheet_name='Sheet1',
                   index=False)  # Id estudante (cod. curso, nome curso, escola, cod. estudante, nome estudante) certificação (Ex: "desenvolvedor de jogos 2d"), período, disciplinas cumpridas

# Close the Pandas Excel writer and output the Excel file.# Concatenar Dataframes
writer.save()

# Dados Essenciais
# Id estudante (cod. curso, nome curso, escola, cod. estudante, nome estudante) certificação (Ex: "desenvolvedor de jogos 2d"), período, disciplinas cumpridas

# Concatenar Dataframes
# frames = [df1, df2, df3]
# result = pd.concat(frames)
