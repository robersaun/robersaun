"""Gerador de Arquivo Unificado de Matrizes

Script para unificar todas as matrizes baixadas do Prime.
Recebe como entrada a pasta raiz contendo todos os arquivos de matrizes em formato XLSX,
e gera um arquivo unificado no formato Excel.
As matrizes do Prime originalmente vêm em formato XLS, para converter em XLSX deve-se abrí-las no Excel e 'Salvar como'.

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

    disciplinas = []
    for dirpath, dirnames, filenames in os.walk(diretorio):
        for file in filenames:
            if "xlsx" in file:
                count += 1
                caminho_completo = dirpath + "/" + file
                caminho_completo = caminho_completo.replace("\\", "/")

                print(f"Lendo arquivo {count}: {caminho_completo.replace(diretorio, '')}")
                try:
                    excel = ler_excel(caminho_completo)
                    for disciplina in excel:
                        disciplinas.append(disciplina)
                except:
                    print("\tNão foi possível ler o arquivo")
    print(f"Total de arquivos: {count}")

    dataframe = pandas.DataFrame(disciplinas)
    dataframe.columns = ['Disciplina', 'Classificação', 'Créditos', 'Grupo', 'CH Teórica', 'CH Prática',
                         'CH Oficial', 'CH Relógio Oficial', 'Curso', 'Matriz', 'Equivalência(s)']
    dataframe.to_excel(excel_writer="excel-unificado.xlsx", index=False)


def ler_excel(arquivo):
    disciplinas_excel = []
    workbook_excel = xlrd.open_workbook(arquivo)
    sheet_excel = workbook_excel.sheet_by_name("Grade Curricular")
    matriz = sheet_excel.cell_value(1, 2)
    curso = sheet_excel.cell_value(2, 2)

    sheet_excel = workbook_excel.sheet_by_name("Ciclos")
    for linha in range(0, sheet_excel.nrows):
        if sheet_excel.cell_value(linha, 1) == "" or \
                sheet_excel.cell_value(linha, 1) == "Total Carga Horária (mínimo):" or \
                sheet_excel.cell_value(linha, 1) == "Disciplina" or \
                sheet_excel.cell_value(linha, 1) == "Ciclo:" or \
                sheet_excel.cell_value(linha, 1) == "Total":
            continue

        nome_disciplina = sheet_excel.cell_value(linha, 1).strip()
        classificacao = sheet_excel.cell_value(linha, 2)
        grupo = sheet_excel.cell_value(linha, 5).strip()

        ch_teorica = sheet_excel.cell_value(linha, 6)
        if ch_teorica == "":
            ch_teorica = 0
        ch_teorica = int(ch_teorica)

        ch_pratica = sheet_excel.cell_value(linha, 7)
        if ch_pratica == "":
            ch_pratica = 0
        ch_pratica = int(ch_pratica)

        ch_oficial = sheet_excel.cell_value(linha, 8)
        if ch_oficial == "":
            ch_oficial = 0
        ch_oficial = int(ch_oficial)

        ch_relogio_oficial = sheet_excel.cell_value(linha, 12)
        if ch_relogio_oficial == "":
            ch_relogio_oficial = 0
        ch_relogio_oficial = int(ch_relogio_oficial)

        creditos = sheet_excel.cell_value(linha, 15)
        if creditos == "":
            creditos = 0
        creditos = int(creditos)

        equivalencia = sheet_excel.cell_value(linha, 18).replace("<br />", "\n")

        disciplinas_excel.append([nome_disciplina, classificacao, creditos, grupo, ch_teorica, ch_pratica, ch_oficial,
                                  ch_relogio_oficial, curso, matriz, equivalencia])

    return disciplinas_excel


if __name__ == '__main__':
    main()
