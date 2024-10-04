# -*- coding: Windows-1252 -*-

from tkinter import Tk, filedialog
import pandas as pd
import os
from datetime import datetime


# Função principal
# noinspection PyBroadException
def main(root):
    pasta = "Unificados Unificado"
    arquivos = []
    itens = []
    colunas = []  # 25 colunas
    # Para cada diretório dentro do diretório raiz
    for dirpath, dirnames, pasta_unificados in os.walk(root):

        # Se o diretório possuir outros diretórios dentro dele, ignora
        if dirnames:
            continue

        # Mostra qual foi o arquivo selecionado
        for nome_arquivo in pasta_unificados:
            if "Excel-Unificado" in nome_arquivo:
                print(f"Arquivo selecionado: {nome_arquivo}")
                arquivos.append(f'{dirpath}/{nome_arquivo}')

                print('> Lendo arquivo')
                df = pd.read_excel(f'{dirpath}/{nome_arquivo}', sheet_name=0)

                if not colunas:
                    header = df.head()
                    colunas.append('Campus')
                    for h in header:
                        colunas.append(h)

                df.fillna('', inplace=True)
                valores = df.values.tolist()
                campus = nome_arquivo.split('-')[2]
                for v in valores:
                    v.insert(0, campus)
                    itens.append(v)
            else:
                pass

    data = datetime.today().strftime('%d-%m-%Y')
    os.makedirs(f'{pasta}', exist_ok=True)
    print('\nGerando excel...')
    dataframe = pd.DataFrame(itens)
    dataframe.columns = colunas
    dataframe.to_excel(excel_writer=f'{pasta}/Unificados-{data}.xlsx', index=False)
    print('Excel gerado')


if __name__ == '__main__':
    # Chamada da função main
    Tk().withdraw()
    pasta_raiz = filedialog.askdirectory()
    main(pasta_raiz)
