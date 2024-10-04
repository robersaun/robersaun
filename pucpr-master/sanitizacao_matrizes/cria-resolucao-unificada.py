"""Gerador de Arquivo Unificado de Resoluções

Script para unificar todas as resoluções.
Recebe como entrada a pasta raiz contendo todos os arquivos de resoluções em formato XLSX,
e gera um arquivo unificado no formato Excel.
As resoluções originalmente vêm em formato Word, e para converter para Excel deve-se copiar todos os dados e colar em
uma nova planilha, cuidando para que cada disciplina fique em uma única linha.

Pode ser utilizado pelo analista de TI para auxiliar no processo de sanitização de matrizes.


Desenvolvido por Vinicius Tozo
Última atualização: 24/08/2021
"""
import os
from tkinter import filedialog, Tk

import pandas
import xlrd


def main():
    Tk().withdraw()
    diretorio = filedialog.askdirectory()

    count = 0
    dataframe_unificado = pandas.DataFrame()

    for dirpath, dirnames, filenames in os.walk(diretorio):
        for file in filenames:

            if file != "Novo(a) Planilha do Microsoft Excel.xlsx":
                continue

            count += 1
            caminho_completo = dirpath + "/" + file
            caminho_completo = caminho_completo.replace("\\", "/")
            dirpath = dirpath.replace("\\", "/")

            print(f"Lendo arquivo {count}: {caminho_completo.replace(diretorio, '')}")
            dataframe_matriz = ler_resolucao(caminho_completo)

            dataframe_matriz["MATRIZ"] = dirpath.split('/')[-1]
            dataframe_matriz["CURSO"] = dirpath.split('/')[-2]
            dataframe_matriz["ESCOLA"] = dirpath.split('/')[-3]
            dataframe_matriz["CAMPUS"] = dirpath.split('/')[-4]

            dataframe_unificado = pandas.concat([dataframe_unificado, dataframe_matriz], axis=0, ignore_index=True)

    print(f"Total de arquivos: {count}")

    dataframe_unificado.sort_values(['CAMPUS', 'ESCOLA', 'CURSO', 'MATRIZ', 'PERÍODO'], inplace=True)
    dataframe_unificado.to_excel(excel_writer="excel-unificado.xlsx", index=False)


def ler_resolucao(arquivo):
    workbook_word = xlrd.open_workbook(arquivo)
    try:
        sheet_word = workbook_word.sheet_by_name("Plan1")
    except:
        sheet_word = workbook_word.sheet_by_name("Planilha1")

    ciclo = 0
    colunas = []
    linhas = []

    for linha in range(0, sheet_word.nrows):

        if type(sheet_word.cell_value(linha, 0)) is str:
            if "PERÍODO" in str.upper(sheet_word.cell_value(linha, 0)):
                ciclo = int(''.join(filter(str.isdigit, sheet_word.cell_value(linha, 0))))
            elif "PERÍODO" in str.upper(str(sheet_word.cell_value(linha, 1))):
                ciclo = int(''.join(filter(str.isdigit, sheet_word.cell_value(linha, 1))))
            elif str.upper(str(sheet_word.cell_value(linha, 1))).strip() in ['DISCIPLINAS', 'AT', 'AP', 'CRED', 'HORAS',
                                                                             'MT', 'MP', 'ORDEM', 'DISCIPLINA',
                                                                             'REQUISITO', 'HA', 'HR']:

                colunas = sheet_word.row_values(linha)

                colunas = [x.upper().strip() for x in colunas]
                colunas = ["ORDEM" if x == "" else x for x in colunas]
                colunas = ["DISCIPLINA" if x == "DISCIPLINAS" else x for x in colunas]
            continue

        valores = {'PERÍODO': ciclo}
        for index, coluna in enumerate(colunas):
            valores[coluna] = str(sheet_word.cell_value(linha, index)).strip()
        linhas.append(valores)

    dataframe_matriz = pandas.DataFrame(linhas)
    return dataframe_matriz


if __name__ == '__main__':
    main()
