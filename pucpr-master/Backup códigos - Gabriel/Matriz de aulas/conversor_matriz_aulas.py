# -*- coding: Windows-1252 -*-

from tkinter import Tk
from tkinter.filedialog import askopenfilename
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

    cont = 1
    
    # Remove os nan das c�lulas vazias
    df_aulas.dropna(inplace=True, how='all') # tira as c�lulas, quando todas estiverem vazias
    df_aulas.drop_duplicates(inplace=True) # tira as c�lulas duplicadas, no caso, os espa�os em branco e os t�tulos
    for n in df_aulas[0]:
        celula = n
        if celula==f'{cont}� PER�ODO':
            df_aulas.drop(columns=0, inplace=True) #Est� pegando o primeiro, mas dps d� erro
            cont += 1
    df_aulas = df_aulas.fillna("")
    
    df_aulas.to_excel('matriz_aulas.xlsx', index=False)

if __name__ == '__main__':
    # Chamada da fun��o main
    main()
