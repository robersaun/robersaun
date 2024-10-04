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

    # Selecionando a planilha PARA
    print("Selecione a planilha PARA")
    Tk().withdraw()
    planilha_para = askopenfilename(
        filetypes=[('Arquivo excel', '.xlsx')],
        title='Selecione a planilha PARA')

    # Mostra qual foi o arquivo selecionado das aulas
    print("Arquivo selecionado: " + str(planilha_aulas).split('/')[-1])
    # Mostra qual foi o arquivo selecionado PARA
    print("Arquivo selecionado: " + str(planilha_para).split('/')[-1])

    ''' -------------------------------------------------------------------------------
        Leitura das p�ginas do arquivo
        ------------------------------------------------------------------------------- '''
    
    print("Lendo arquivo")
    
    # Planilha de aulas � armazenado na vari�vel
    df_aulas = pd.read_excel(planilha_aulas, dtype=str, header=1)
    # planilha PARA e armazenada em uma variavel, selecionando apartir do header 12
    df_para = pd.read_excel(planilha_para, dtype=str, header=12, index_col=1)

    #TODO ver se isso fica assim mesmo:
    # deletando a coluna que n�o vamos usar
    df_para.drop('Unnamed: 0',inplace=True,axis=1)

    df_aulas.insert(loc=0, column='Periodo', value=nan) # Cria uma coluna para colocar o n�mero do per�odo
    cont = 1 # Contador do per�odo
    linha = 0 # Contador de linha
    # Insere o n�mero do per�odo na coluna criada
    # Tira os valores desnecess�rios das linhas
    nome_coluna = [] # Cria uma lista contendo o nome de cada coluna
    for c in df_aulas.columns:
        nome_coluna.append(c)
    for dado in df_aulas['Ordem']:
        if (dado == f'{cont+1}� PER�ODO'):
            cont+=1
        if (str(dado)!='nan'):
            df_aulas['Periodo'][linha] = cont
            if (str(dado)=='Ordem' or str(dado)=='TOTAL' or str(dado)==f'{cont}� PER�ODO'):
                df_aulas['Periodo'][linha] = nan
                coluna = 0
                while coluna < 11:
                    df_aulas[nome_coluna[coluna]][linha] = nan
                    coluna+=1
        linha+=1 
    # Remove os nan das c�lulas vazias
    df_aulas.dropna(inplace=True, how='all') # remove a linha, quando todas as c�lulas estiverem vazias
    df_aulas.drop_duplicates(inplace=True) # remove duplicadas, no caso, os espa�os em branco e os t�tulos
    
    df_aulas = df_aulas.fillna("")


    df_aulas.to_excel('matriz_aulas.xlsx', index=False)

if __name__ == '__main__':
    # Chamada da fun��o main
    main()

