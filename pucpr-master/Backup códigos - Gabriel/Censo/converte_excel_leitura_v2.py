# -*- coding: Windows-1252 -*-

from tkinter import Tk
from tkinter.filedialog import askopenfilename
import pandas as pd

# Fun��o principal
def main():
    # Sele��o de arquivo dos alunos
    print("Selecione a planilha de alunos")
    Tk().withdraw()
    planilha_dados = askopenfilename(
        filetypes=[('Arquivo excel', '.xlsx')],
        title='Selecione a planilha de alunos')

    print("Arquivo selecionado: " + str(planilha_dados).split('/')[-1])

    ''' -------------------------------------------------------------------------------
        Leitura das p�ginas do arquivo
        ------------------------------------------------------------------------------- '''
    
    print("Lendo arquivo")

    df = pd.read_excel(planilha_dados, dtype=str, sheet_name=0)

    dados_alunos = []
    dados_vinculos = []
    
    for i in range(len(df)):
        if i < 86:
            c = df.columns[i]
        if i < 22:
            dados_alunos.append(str(c))
            if c == 'CPF':
                dados_vinculos.append(str(c))
        elif i < 86:
            dados_vinculos.append(str(c))
        else:
            break

    print(".")

    df.dropna
    df = df.fillna("")

    print(".")

    with pd.ExcelWriter('Base_Alunos.xlsx') as writer:  
        df.to_excel(writer, sheet_name='Dados pessoais', index=False, columns=dados_alunos)
        df.to_excel(writer, sheet_name='Vinculos', index=False, columns=dados_vinculos)

    exit()

if __name__ == '__main__':
    # Chamada da fun��o main
    main()
