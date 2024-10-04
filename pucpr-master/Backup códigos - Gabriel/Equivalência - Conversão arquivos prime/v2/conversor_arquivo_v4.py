# -*- coding: Windows-1252 -*-

from tkinter import Tk, filedialog
import pandas as pd
import os

# Used to save the file as excel workbook
# Need to install this library
from xlwt import Workbook
# Used to open to corrupt excel file
import io


# Fun��o principal
# noinspection PyBroadException
def main(root):
    # Para cada diret�rio dentro do diret�rio raiz
    for dirpath, dirnames, planilha in os.walk(root):

        # Se o diret�rio possuir outros diret�rios dentro dele, ignora
        if dirnames:
            continue
        # Sele��o do arquivo

        curso = dirpath.replace('\\', '/').split('/')[-1]

        # nome_arquivo = str(planilha).split('/')[-1]
        # Mostra qual foi o arquivo selecionado
        for nome_arquivo in planilha:
            #try:
                print(f"Arquivo selecionado: {nome_arquivo}")
                doc_nome = nome_arquivo.split('.')[0]
                doc_grade = f"{nome_arquivo.split('_')[-2]}-{nome_arquivo.split('_')[-1].split('.')[0]}"
                os.makedirs(f'{curso}/{doc_grade}', exist_ok=True)

                ''' -------------------------------------------------------------------------------
                    Leitura das p�ginas do arquivo
                    ------------------------------------------------------------------------------- '''

                # Planilha de equival�ncias � armazenada na vari�vel
                # Os dados dos alunos dever�o sempre estar na primeira p�gina do arquivo
                arquivo = f'{dirpath}/{nome_arquivo}'
                df_f = pd.ExcelFile(nome_arquivo).parse(sheet_name=None)
                print(df_f)
                continue

                sheet_names = df_f.sheet_names


                dfn = []
                for s in sheet_names:
                    dfn.append(pd.read_excel(arquivo, dtype=str, sheet_name=s))

                with pd.ExcelWriter(f'{curso}/{doc_grade}/{doc_nome}.xlsx') as writer:
                    for c in range(len(sheet_names)):
                        sn = sheet_names[c]
                        '''header = []
                        for h in dfn[c].head():
                            header.append('')'''
                        if 'Grup ' in sn:
                            sn = 'Grup Atv Comp - Proj Com'
                        dfn[c].to_excel(writer, index=False, header=False, sheet_name=sn, startrow=1)
                print('Excel gerado')
            #except:
            #    pass


if __name__ == '__main__':
    # Chamada da fun��o main
    Tk().withdraw()
    pasta_raiz = filedialog.askdirectory()
    main(pasta_raiz)

#TODO: Add o Ciclo
#TODO: Arrumar para pegar os valores direito - OK? Oqq eu devia fazer aqui?
#TODO: Alterar a gera��o do nome do arquivo - OK
#TODO: Como que vou comparar as cargas hor�rias das equival�ncias?

# Aparentemnte agr � s� selecionar o diret�rio raiz q roda td