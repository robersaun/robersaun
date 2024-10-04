
import json
import random
import re
import string
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from xml.etree import ElementTree
import pandas
from datetime import datetime


def gerarRelatorio(file_name_xlsx):

    df = pandas.read_excel(file_name_xlsx,dtype=str,sheet_name=0)

    #tranformando df em dict para ser mais facil de manipular
    dict_dados ={}
    dict_nome_aluno_situacao_data={}
    key = 0

    for dados in df.iterrows():

        dict_dados[key] = {
            'ESCOLA' : dados[1]['ESCOLA'],
            'RA' : dados[1]['RA'],
            'NOME_ALUNO' : dados[1]['NOME_ALUNO'],
            'SIT_ACADEMICA' : dados[1]['SIT_ACADEMICA'],
            'DT_SIT_ACADEMICA' : dados[1]['DT_SIT_ACADEMICA']
        }

        key+=1
        valor = f"{dados[1]['SIT_ACADEMICA']},{dados[1]['DT_SIT_ACADEMICA']}"

        try:

                dict_nome_aluno_situacao_data[dados[1]['NOME_ALUNO']] +=f"/{valor}"

        except:
            dict_nome_aluno_situacao_data[dados[1]['NOME_ALUNO']] = f"{valor}"

    dict_nome_diferanca ={}
    for dados in dict_nome_aluno_situacao_data.items():
        quantidade = dados[1].count('/')
        if quantidade == 0:
            continue
        data_anterior = None
        diferenca_total = 0
        #se for numero impar é um tracamento
        verifica_impar = 1

        for posicao in range(0,quantidade+1):
            if verifica_impar %2 != 0:
                data_anterior = dados[1].split('/')[posicao].split(',')[1].split(' ')[0]
                verifica_impar +=1
                continue
            data_proximo = dados[1].split('/')[posicao].split(',')[1].split(' ')[0]

            d1 = datetime.strptime(data_anterior, "%Y-%m-%d")
            d2 = datetime.strptime(data_proximo, "%Y-%m-%d")
            if d1 == d2:
                continue
            try:
                diferenca = d2 - d1
                diferenca_total += int(str(diferenca).split(' ')[0])
                verifica_impar += 1
            except:
                print(f"{dados[0]} {d1} {d2}")
        dict_nome_diferanca[dados[0]] = diferenca_total
    print('')
def main():

   #selecionando arquivos
    print('Selecione o relatório para tratamento dos vinculos')
    Tk().withdraw()
    file_name_xlsx = askopenfilename(filetypes=[('xlsx', '.xlsx')])

    gerarRelatorio(file_name_xlsx)

if __name__ == '__main__':
    main()
