'''
Criado para listar todos os documentos de uma pasta selecionada pelo usuário.

Desenvolvido por Vinicius Tozo
última atualização: 25/01/2021
'''

import os
from tkinter import Tk, filedialog

import pandas


def lista_documentos(root):
    # Lista todos os arquivos dentro de uma pasta escolhida pelo usuário
    # Criado para listar documentos dos professores na pasta do sharepoint

    lista = []

    # Para cada diretório dentro do diretório raiz
    for dirpath, dirnames, filenames in os.walk(root):

        # Para cada arquivo dentro do diretório, verifica se a matriz possui os arquivos necessários
        for file in filenames:
            lista.append([dirpath.split('/')[-1], file])

    df = pandas.DataFrame(lista, columns=['Pasta', 'Arquivo'])
    df.to_excel('Lista de documentos.xlsx', index=False)


if __name__ == '__main__':
    Tk().withdraw()
    print('Selecione a pasta de documentos')
    pasta_raiz = filedialog.askdirectory()
    lista_documentos(pasta_raiz)
    print('Arquivo "Lista de documentos.xlsx" gerado')
    input('Pressione enter para sair...')
