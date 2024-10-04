# -*- coding: Windows-1252 -*-

from tkinter import Tk
from tkinter.filedialog import askopenfilename
from numpy import nan
import pandas as pd

# Fun��o principal
def main():
    # Sele��o de arquivo dos aulas
    print("Selecione a planilha de aulas")
    Tk().withdraw()
    planilha_aulas = askopenfilename(
        filetypes=[('Arquivo excel', '.xlsx')],
        title='Selecione a planilha de aulas')

    # Mostra qual foi o arquivo selecionado
    print("Arquivo selecionado: " + str(planilha_aulas).split('/')[-1])

    ''' -------------------------------------------------------------------------------
        Leitura das p�ginas do arquivo
        ------------------------------------------------------------------------------- '''
    
    print("Lendo arquivo")
    
    # Planilha de aulas � armazenado na vari�vel
    df_aulas = pd.read_excel(planilha_aulas, dtype=str, header=None)

    cont = 1 # Contador do per�odo
    df_aulas.insert(loc=0, column='Periodo', value=nan) # Cria uma coluna para colocar o n�mero do per�odo
    linha = 0 # Cantador de linha
    for dado in df_aulas[0]:
        if (dado == f'{cont+1}� PER�ODO'):
            cont+=1
        if (str(dado)!='nan'):
            df_aulas['Periodo'][linha] = cont
        linha+=1 
    # Remove os nan das c�lulas vazias
    df_aulas.dropna(inplace=True, how='all') # remove a linha, quando todas as c�lulas estiverem vazias
    df_aulas.drop_duplicates(inplace=True) # remove duplicadas, no caso, os espa�os em branco e os t�tulos
    
    '''qtd_linhas = df_aulas.shape[1]
    index_linhas = []
    n = 0
    while n != qtd_linhas:
        index_linhas.append(n)
        n+=1'''
    '''for n in df_aulas[0]:
        celula = n
        if celula==f'{cont}� PER�ODO':
            df_aulas.drop(columns=0, inplace=True) #Est� pegando o primeiro, mas dps d� erro'''
    cont += 1
    df_aulas = df_aulas.fillna("")
    
    df_aulas.to_excel('matriz_aulas.xlsx', index=False)

if __name__ == '__main__':
    # Chamada da fun��o main
    main()
