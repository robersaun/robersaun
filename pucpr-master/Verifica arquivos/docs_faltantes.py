'''
Lista os documetos dos professores presentes na pasta do sherepoint e indica os documentos faltantes. Ele recebe uma pasta de documentos para realizar a verificação.

Desenvolvido por Gabriel Ernesto
Última atualização: 26/01/2021
'''

import os
from tkinter import Tk, filedialog
from tkinter.filedialog import askopenfilename
import pandas as pd
from dateutil.utils import today


def lista_documentos():
    # Lista todos os arquivos dentro de uma pasta escolhida pelo usuário
    # Criado para listar documentos dos professores na pasta do sharepoint

    lista = askopenfilename(
        filetypes=[('Arquivo excel', '.xlsx')],
        title='Selecione a planilha de documentos')

    df = pd.read_excel(lista, usecols='A:B')

    # Nome do prof e tipo de arquivo
    prof = []
    tipo = []
    docs = []
    pend = []
    for i in range(len(df['0'])):
        # Dados prof
        nome_prof = str(df['0'][i].replace('\\', '/').split('/')[-2]).upper()
        if len(nome_prof) == 1:
            nome_prof = str(df['0'][i].replace('\\', '/').split('/')[-1]).upper()
        prof.append(nome_prof)

        # Nome dos documentos
        nome_doc = str(df['1'][i].replace('-', '_')).upper()
        # Tira o indicativo de que é um pdf
        nome_doc = nome_doc.split('.pdf')[0]
        nome_doc = nome_doc.split('.PDF')[0]
        # Tira o nome do prof, para deixar só a data ou o tipo do doc
        # Para clareza no excel
        nome_doc = nome_doc.split(f'{nome_prof}_')[-1]
        nome_doc = nome_doc.split(f'{nome_prof} _')[-1]
        nome_doc = nome_doc.split(f'{nome_prof}_ ')[-1]
        nome_doc = nome_doc.split(f'{nome_prof} ')[-1]
        docs.append(nome_doc)

        # Tipo de documento
        tipo_doc = str(df['0'][i].replace('\\', '/').split('/')[-1]).upper()
        if 'TACH' in nome_doc:
            tipo_doc = 'TACH'
            tipo.append('TACH')
        elif 'CPCD' in nome_doc:
            tipo_doc = 'CPCD'
            tipo.append('CPCD')
        elif 'NPCD' in nome_doc:
            tipo_doc = 'NPCD'
            tipo.append('NPCD')
        elif nome_prof == tipo_doc:
            tipo.append('Desconhecido')
        else:
            tipo.append(tipo_doc)

        # Ajuste dos dados
        if i > 1:
            n = -2
        else:
            n = -1
        td = tipo[n]
        np = prof[n]
        nd = docs[n]

        if nome_prof == np and tipo_doc == td and nome_doc != nd:
            nome_doc = nd + f', {nome_doc}'
            docs[-2] = nome_doc
            docs.pop()
            tipo.pop()
            prof.pop()

    # Identificar os documentos que faltam
    ano = int(today().year)

    for p in range(len(tipo)):
        # Considerando a data inicial 2010
        data = int(2010)
        pendencia = ''
        # Nem todos os tipos de documentos tem datas
        if tipo[p] == 'TACH' or tipo[p] == 'CPCD':
            while data < ano:
                if tipo[p] == 'CPCD':
                    if str(data) not in docs[p]:
                        if pendencia == '':
                            pendencia = f'{data}'
                        else:
                            pendencia = pendencia + f', {data}'
                elif tipo[p] == 'TACH':
                    if f'{data}.1' not in docs[p]:
                        if pendencia == '':
                            pendencia = f'{data}.1'
                        else:
                            pendencia = pendencia + f', {data}.1'
                    if f'{data}.2' not in docs[p]:
                        if pendencia == '':
                            pendencia = f'{data}.2'
                        else:
                            pendencia = pendencia + f', {data}.2'
                data += 1
        pend.append(pendencia)

    ndf = pd.DataFrame(data={'Professores': prof,
                             'Tipo de arquivo': tipo,
                             'Documentos existentes': docs,
                             'Pendências': pend})
    n_arquivo = 'Documentos Faltantes 10.xlsx'
    ndf.to_excel(n_arquivo, index=False)
    print(f'Arquivo "{n_arquivo}" gerado')


if __name__ == '__main__':
    Tk().withdraw()
    print('Selecione o arquivo')
    lista_documentos()

# TODO: Ver qual a data de corte para considerar os documentos
# TODO: Ver quais são os tipos de documentos que possuem data
