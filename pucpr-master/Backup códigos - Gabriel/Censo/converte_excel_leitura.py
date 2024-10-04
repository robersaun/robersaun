# -*- coding: Windows-1252 -*-

# TODO: Ajustar isso para criar as planilhas automaticamente no formato certo para a leitura no censo

from tkinter import Tk
from tkinter.filedialog import askopenfilename
import pandas as pd

# Fun��o principal
def main():
    # Sele��o de arquivo dos alunos
    print("Selecione a planilha de alunos")
    Tk().withdraw()
    planilha_alunos = askopenfilename(
        filetypes=[('Arquivo excel', '.xlsx')],
        title='Selecione a planilha de alunos')

    print("Arquivo selecionado: " + str(planilha_alunos).split('/')[-1])

    ''' -------------------------------------------------------------------------------
        Leitura das p�ginas do arquivo
        ------------------------------------------------------------------------------- '''
    
    print("Lendo arquivo")

    alunos_escolas = pd.ExcelFile(planilha_alunos)
    nome_escolas = []
    # Verifica quais s�o os nomes das p�ginas para cada escola
    # Obrigat�rio que a primeira seja a do v�nculo
    for escola in range(len(alunos_escolas.sheet_names)):
        if escola != 0:
            nome_escolas.append(escola)

    print(".")

    # Planilha de alunos � armazenado na vari�vel conforme a escola 
    df_alunos = pd.read_excel(planilha_alunos, dtype=str, sheet_name=nome_escolas)
    # Pega as informa��es da p�gina de v�nculos
    df_vinculos = pd.read_excel(planilha_alunos, dtype=str, sheet_name=0)
    
    print(".")

    # A princ�pio tem que tratar as linhas que s� tem as informa��es de totais
    # Concatena as p�ginas, em uma �nica
    # Obrigat�rio que todas as colunas que contenham as mesmas informa��es possuam o mesmo nome
    # Se os dados ao fim da tabela forem indesejados, remover antes de fazer a convers�o
    df_alunos = pd.concat(df_alunos, ignore_index=True)

    print(".")

    df_alunos.insert(loc=0, column="Tipo de registro", value="41")
    df_vinculos.pop(df_vinculos.columns[0])
    df_vinculos.insert(loc=1, column="Tipo de registro", value="42")

    df_alunos.dropna()

    # Remove os nan das c�lulas vazias
    df_alunos = df_alunos.fillna("")
    df_vinculos = df_vinculos.fillna("")

    with pd.ExcelWriter('Base Alunos.xlsx') as writer:  
        df_alunos.to_excel(writer, sheet_name='Dados pessoais', index=False)
        df_alunos.pop(df_vinculos.columns[1])
        df_vinculos.to_excel(writer, sheet_name='Vinculos', index=False)
        # Os dados desnecess�rios dever�o ser removidos manualmente

    exit()

if __name__ == '__main__':
    # Chamada da fun��o main
    main()
