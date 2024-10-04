'''
Script faz um tratamento dos dados da tabela excel Protocolos, retirando protocolos duplicados e datas vazias

Desenvolvido por Vinicius Tozo
Última atualização: 29/01/2021
'''

import pandas
from tkinter import Tk
from tkinter.filedialog import askopenfilename


def main():
    Tk().withdraw()
    arquivo = askopenfilename(
        filetypes=[('Arquivo excel', '.xlsx')],
        title='Selecione o relatório de protocolos do Prime')

    dataframe = pandas.read_excel(arquivo)
    dataframe = dataframe.drop(dataframe[dataframe['Data Evento'] == '00/00/00'].index)
    dataframe = dataframe.drop_duplicates('Nº Prot', keep='last')
    dataframe.to_excel('ProtocolosPrime.xlsx', index=False, header=True)


if __name__ == '__main__':
    main()
