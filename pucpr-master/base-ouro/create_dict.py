"""Script para criação de dicionários

Script para auxiliar na criação de dicionários Curso-CR a partir da Base Ouro.
Recebe como entrada a Base Ouro em formato XLSX.
Gera dicionário em formato TXT. É preciso preencher valores vazios manualmente antes de utilizar o dicionário.

Pode ser utilizado pelo analista de TI para atualizar dicionários de Curso-CR, incluindo no script da Base Ouro
"""
import sys
from tkinter import Tk
from tkinter.filedialog import askopenfilename

import pandas


def encontra_curso(turma):
    if not turma.__contains__("-"):
        return ""
    curso = turma.split("-")[0].strip()
    return curso


def encontra_turno(turma):
    if not turma.__contains__("-"):
        return ""
    turno = turma.split("-")[2].strip()
    return turno


def create_dict_turma_curso(base_ouro):
    df = pandas.read_excel(base_ouro)
    dict_cursos = {}
    dict_txt = open("dict.txt", "w")
    sys.stdout = dict_txt

    for index, linha in df.iterrows():
        cod_curso = encontra_curso(linha["Turma Aluno"])
        nome_curso = linha["Curso Aluno"]

        chave = cod_curso
        if chave not in dict_cursos:
            dict_cursos[chave] = nome_curso

    for key, value in dict_cursos.items():
        print("'" + str(key) + "': '" + str(value) + "',")


def create_dict_curso_cr(base_ouro):
    df = pandas.read_excel(base_ouro)
    dict_cursos = {}
    dict_txt = open("dict.txt", "w")
    sys.stdout = dict_txt

    for index, linha in df.iterrows():
        cod_curso = encontra_curso(linha["Turma Aluno"])
        turno = encontra_turno(linha["Turma Aluno"])
        cr_curso = linha["CR Aluno"]

        chave = cod_curso + " " + turno
        if chave not in dict_cursos:
            dict_cursos[chave] = cr_curso

    for key, value in dict_cursos.items():
        print("'" + str(key) + "': '" + str(value) + "',")


if __name__ == '__main__':
    Tk().withdraw()
    caminho_base_ouro = askopenfilename(
        filetypes=[('Arquivo excel', '.xlsx')],
        title='Selecione a Base Ouro')
    create_dict_curso_cr(caminho_base_ouro)
