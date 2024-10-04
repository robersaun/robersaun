'''
Converte o arquivo do XML do Urania para um arquivo em Excel

Desenvolvido por Vinicius Tozo
Última atualização: 16/07/2021
'''
import xml.etree.ElementTree as ElementTree
from tkinter import Tk
from tkinter.filedialog import askopenfilename

import xlsxwriter


def recebe_nome_do_arquivo():
    Tk().withdraw()
    file_name = askopenfilename(filetypes=[('XML', '.xml')])
    return file_name


def main():
    print("Selecione o arquivo XML")
    file_name_xml = recebe_nome_do_arquivo()
    file_name = file_name_xml.replace(".XML", "").replace(".xml", "")

    tree_xml = ElementTree.parse(file_name_xml)
    root = tree_xml.getroot()

    # Cria um arquivo excel e adiciona uma aba para cada informação
    workbook = xlsxwriter.Workbook(f'{file_name}.xlsx')
    for aba in root:
        worksheet = workbook.add_worksheet(aba.tag)
        for row, registro in enumerate(aba):
            for col, atributo in enumerate(registro):
                if row == 0:
                    worksheet.write(row, col, atributo.tag)
                worksheet.write(row + 1, col, atributo.text)

    workbook.close()
    print("Arquivo gerado com sucesso!")


if __name__ == "__main__":
    main()
