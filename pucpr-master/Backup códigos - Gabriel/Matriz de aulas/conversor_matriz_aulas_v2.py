# -*- coding: Windows-1252 -*-

from tkinter import Tk
from tkinter.filedialog import askopenfilename
from numpy import nan
import pandas as pd

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
    df_aulas = pd.read_excel(planilha_aulas, dtype=str, header=None)

    cont = 1 # Contador do período
    df_aulas.insert(loc=0, column='Periodo', value=nan) # Cria uma coluna para colocar o número do período
    linha = 0 # Cantador de linha
    for dado in df_aulas[0]:
        if (dado == f'{cont+1}º PERÍODO'):
            cont+=1
        if (str(dado)!='nan'):
            df_aulas['Periodo'][linha] = cont
        linha+=1 
    # Remove os nan das células vazias
    df_aulas.dropna(inplace=True, how='all') # remove a linha, quando todas as células estiverem vazias
    df_aulas.drop_duplicates(inplace=True) # remove duplicadas, no caso, os espaços em branco e os títulos
    
    '''qtd_linhas = df_aulas.shape[1]
    index_linhas = []
    n = 0
    while n != qtd_linhas:
        index_linhas.append(n)
        n+=1'''
    '''for n in df_aulas[0]:
        celula = n
        if celula==f'{cont}º PERÍODO':
            df_aulas.drop(columns=0, inplace=True) #Está pegando o primeiro, mas dps dá erro'''
    cont += 1
    df_aulas = df_aulas.fillna("")
    
    df_aulas.to_excel('matriz_aulas.xlsx', index=False)

if __name__ == '__main__':
    # Chamada da função main
    main()
