"""Unifica relatórios diários de ajuste acadêmico

Script para extrair a descrição de todos os protocolos a partir de uma pasta
contendo os relatórios de ajuste acadêmico, e gerando um arquivo em excel com
a lista de descrições.

Não é mais utilizado, projeto descontinuado.
"""
import os
import pandas
import xlrd
import re
from tkinter import filedialog
from tkinter import Tk


def unificar_relatorios(root):
    contagem = 0
    dataframe = pandas.DataFrame(columns=["Nº Protocolo", "Comentários"])

    # Para cada diretório dentro do diretório raiz
    for dirpath, dirnames, filenames in os.walk(root):

        # Para cada arquivo dentro do diretório, verifica se a past possui os arquivos necessários
        for file in filenames:
            if "relatório periódico" in file and "xlsx" in file:
                contagem += 1
                caminho_completo = dirpath + "/" + file
                caminho_completo = caminho_completo.replace("\\", "/")

                print(f"Lendo arquivo {contagem}: {caminho_completo}")
                try:
                    excel = pandas.read_excel(caminho_completo, sheet_name="Caixa de Trabalho")
                    dataframe = pandas.concat([dataframe, excel])
                except:
                    print("Não foi possível extrair os dados")
    return dataframe


def main():
    Tk().withdraw()
    pasta_raiz = filedialog.askdirectory()
    relatorio_unificado = unificar_relatorios(pasta_raiz)
    relatorio_unificado.to_excel("relatorio_unificado.xlsx")


if __name__ == '__main__':
    main()
