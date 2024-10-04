# -*- coding: Windows-1252 -*-

from tkinter import Tk
from tkinter.filedialog import askopenfilename
from numpy import nan
import pandas as pd

# TODO: Acho q da para tentar pegar a tabela PARA, para escrever nela, os valores conforme o necessário

# Função principal
def main():
    # Seleção de arquivo dos aulas
    print("Selecione a planilha de aulas")
    Tk().withdraw()
    planilha_aulas = askopenfilename(
        filetypes=[('Arquivo excel', '.xlsx')],
        title='Selecione a planilha de aulas')

    # Mostra qual foi o arquivo selecionado
    print("Arquivo selecionado: " + str(planilha_aulas).split('/')[-1])

    ''' -------------------------------------------------------------------------------
        Leitura das páginas do arquivo
        ------------------------------------------------------------------------------- '''
    
    print("Lendo arquivo")
    
    # Planilha de aulas é armazenado na variável
    df_aulas = pd.read_excel(planilha_aulas, dtype=str, header=1)

    df_aulas.insert(loc=0, column='Periodo', value=nan) # Cria uma coluna para colocar o número do período
    cont = 1 # Contador do período
    linha = 0 # Contador de linha
    # Insere o número do período na coluna criada
    # Tira os valores desnecessários das linhas
    nome_coluna = [] # Cria uma lista contendo o nome de cada coluna
    for c in df_aulas.columns:
        nome_coluna.append(c)
    for dado in df_aulas['Ordem']:
        if (dado == f'{cont+1}º PERÍODO'):
            cont+=1
        if (str(dado)!='nan'):
            df_aulas['Periodo'][linha] = cont
            if (str(dado)=='Ordem' or str(dado)=='TOTAL' or str(dado)==f'{cont}º PERÍODO'):
                df_aulas['Periodo'][linha] = nan
                coluna = 0
                while coluna < 11:
                    df_aulas[nome_coluna[coluna]][linha] = nan
                    coluna+=1
        linha+=1 
    # Remove os nan das células vazias
    df_aulas.dropna(inplace=True, how='all') # remove a linha, quando todas as células estiverem vazias
    df_aulas.drop_duplicates(inplace=True) # remove duplicadas, no caso, os espaços em branco e os títulos
    
    df_aulas = df_aulas.fillna("")
    
    df_aulas.to_excel('matriz_aulas.xlsx', index=False)

if __name__ == '__main__':
    # Chamada da função main
    main()
