'''
Script recebe um diretório de arquivos das turmas e gera um excel com a unificação dessas turmas

Desenvolvido por Vinicius Tozo
Útlima atualização: 23/06/2021
'''

import os
from tkinter import Tk, filedialog

import pandas


def unifica_turmas(root):
    colunas_global_class = ["ESCOLA", "Cód. Curso", "Curso ofertante", "Período da disciplina", "Tipo de disciplina (Obrigatória, optativa ou eletiva)", "Nível 1, 2 ou 3 (Global Classes)", "Turma", "A disciplina pode ser ofertada para outros curso?", "Código da disciplina", "Nome da disciplina", "Código da disciplina", "Nome da disciplina", "Professor", "Nº de Vagas", "Dia", "Início", "Fim", "CH TOTAL DA DISCIPLINA", "Hora aula total da disciplina", "Observação", "E-mail Institucional"]

    df_regulares = pandas.DataFrame()
    df_eletivas = pandas.DataFrame()
    df_especiais = pandas.DataFrame()
    df_global_classes = pandas.DataFrame(columns=colunas_global_class)

    # Para cada diretório dentro do diretório raiz
    for dirpath, dirnames, filenames in os.walk(root):

        # Para cada arquivo dentro do diretório
        for file in filenames:

            if not file.endswith(".xlsx"):
                continue

            regulares = pandas.read_excel(
                dirpath + "/" + file, sheet_name="Disciplinas Regulares", skiprows=12, engine="openpyxl")
            eletivas = pandas.read_excel(
                dirpath + "/" + file, sheet_name="Eletivas Próprias", skiprows=11, engine="openpyxl")
            especiais = pandas.read_excel(
                dirpath + "/" + file, sheet_name="Turmas Especiais Previstas", skiprows=8, engine="openpyxl")
            global_classes = pandas.read_excel(
                dirpath + "/" + file, sheet_name="Global Classes", skiprows=7, engine="openpyxl")
            global_classes.columns = colunas_global_class

            df_regulares = pandas.concat([df_regulares, regulares])
            df_eletivas = pandas.concat([df_eletivas, eletivas])
            df_especiais = pandas.concat([df_especiais, especiais])
            df_global_classes = pandas.concat([df_global_classes, global_classes])

    df_regulares.dropna(how='all', inplace=True)
    df_eletivas.dropna(how='all', inplace=True)
    df_especiais.dropna(how='all', inplace=True)
    df_global_classes.dropna(how='all', inplace=True)

    writer = pandas.ExcelWriter('Unificado.xlsx', engine='xlsxwriter')

    df_regulares.to_excel(writer, index=False, sheet_name="Disciplinas Regulares")
    df_eletivas.to_excel(writer, index=False, sheet_name="Eletivas Próprias")
    df_especiais.to_excel(writer, index=False, sheet_name="Turmas Especiais Previstas")
    df_global_classes.to_excel(writer, index=False, sheet_name="Global Classes")

    writer.save()


def main():
    Tk().withdraw()
    pasta_raiz = filedialog.askdirectory()
    unifica_turmas(pasta_raiz)
    input("Arquivo gerado com sucesso!")


if __name__ == '__main__':
    main()
