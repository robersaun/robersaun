# -*- coding: Windows-1252 -*-

# TODO: Ajustar isso para criar as planilhas automaticamente no formato certo para a leitura no censo

from tkinter import Tk
from tkinter.filedialog import askopenfilename
import pandas as pd

# Função principal
def main():
    # Seleção de arquivo dos alunos
    print("Selecione a planilha de alunos")
    Tk().withdraw()
    planilha_alunos = askopenfilename(
        filetypes=[('Arquivo excel', '.xlsx')],
        title='Selecione a planilha de alunos')

    print("Arquivo selecionado: " + str(planilha_alunos).split('/')[-1])

    ''' -------------------------------------------------------------------------------
        Leitura das páginas do arquivo
        ------------------------------------------------------------------------------- '''
    
    print("Lendo arquivo")

    alunos_escolas = pd.ExcelFile(planilha_alunos)
    nome_escolas = []
    # Verifica quais são os nomes das páginas para cada escola
    # Obrigatório que a primeira seja a do vínculo
    for escola in range(len(alunos_escolas.sheet_names)):
        if escola != 0:
            nome_escolas.append(escola)

    print(".")

    # Planilha de alunos é armazenado na variável conforme a escola 
    df_alunos = pd.read_excel(planilha_alunos, dtype=str, sheet_name=nome_escolas)
    # Pega as informações da página de vínculos
    df_vinculos = pd.read_excel(planilha_alunos, dtype=str, sheet_name=0)
    
    print(".")

    # A princípio tem que tratar as linhas que só tem as informações de totais
    # Concatena as páginas, em uma única
    # Obrigatório que todas as colunas que contenham as mesmas informações possuam o mesmo nome
    # Se os dados ao fim da tabela forem indesejados, remover antes de fazer a conversão
    df_alunos = pd.concat(df_alunos, ignore_index=True)

    print(".")

    df_alunos.insert(loc=0, column="Tipo de registro", value="41")
    df_vinculos.pop(df_vinculos.columns[0])
    df_vinculos.insert(loc=1, column="Tipo de registro", value="42")

    df_alunos.dropna()

    # Remove os nan das células vazias
    df_alunos = df_alunos.fillna("")
    df_vinculos = df_vinculos.fillna("")

    with pd.ExcelWriter('Base Alunos.xlsx') as writer:  
        df_alunos.to_excel(writer, sheet_name='Dados pessoais', index=False)
        df_alunos.pop(df_vinculos.columns[1])
        df_vinculos.to_excel(writer, sheet_name='Vinculos', index=False)
        # Os dados desnecessários deverão ser removidos manualmente

    exit()

if __name__ == '__main__':
    # Chamada da função main
    main()
