"""Gerador de Pastas para as Matrizes

Script para gerar a estrutura de pastas utilizada na sanitização de matrizes.
Recebe como entrada um arquivo no formato Excel contendo quatro colunas: 'Campus', 'Escola', 'Curso' e 'Matriz'.

Gera as pastas aninhadas, na mesma ordem:
    Campus
        |- Escola
            |- Curso
                |- Matriz

Exemplo:
    Curitiba
        |- Belas Artes
            |- Arquitetura e Urbanismo
                |- CAUR 2022-1

Pode ser utilizado pelo analista de TI para auxiliar no processo de sanitização de matrizes.


Desenvolvido por Vinicius Tozo
Última atualização: 19/10/2021
"""
import os
from tkinter import Tk
from tkinter.filedialog import askopenfilename

import pandas


def trata_caracteres(string):
    return string.replace("/", "-").replace(":", "")


def main():
    diretorio_pai = ""

    Tk().withdraw()
    excel_path = askopenfilename(
        filetypes=[('Arquivo excel', '.xlsx')],
        title='Selecione o arquivo com as informações do grupo das disciplinas')
    df = pandas.read_excel(excel_path)

    for index, linha in df.iterrows():
        diretorio = (
                trata_caracteres(linha["Campus"]) + '/' +
                trata_caracteres(linha["Escola"]) + '/' +
                trata_caracteres(linha["Curso"]) + '/' +
                trata_caracteres(linha["Matriz"]) + '/'
        )
        path = os.path.join(diretorio_pai, diretorio)
        try:
            os.makedirs(path)
        except OSError:
            print("Directory '%s' can not be created" % path)


if __name__ == '__main__':
    main()
