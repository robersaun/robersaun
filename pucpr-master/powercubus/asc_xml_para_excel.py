"""Conversor do XML do aSc para o formato Excel

Script para gerar um arquivo excel com os dados extraídos do aSc via 'exportação XML para Prime'.
Recebe como entrada o arquivo xml gerado pelo script aSc.
Não realiza modificações nos dados, apenas converte o formato.
Por isso pode ser utilizado para outros arquivos XML se necessário.

Utilizado pelo analista de TI no processo de elaboração de horários no PowerCubus.

Desenvolvido por Vinicius Tozo
Última atualização: 06/01/2022
"""
import xml.etree.ElementTree as ElementTree
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import xlsxwriter


def recebe_nome_do_arquivo():
    Tk().withdraw()
    file_name = askopenfilename(filetypes=[('XML', '.xml')])
    return file_name


def corrige_caracteres_especiais(file_name):
    # Lê o arquivo
    with open(file_name, 'r', encoding='windows-1252') as file:
        file_data = file.read()

    # Substitui o símbolo que causa erro
    file_data = file_data.replace(' & ', ' &amp; ')

    # Sobrescreve o arquivo
    with open(file_name, 'w') as file:
        file.write(file_data)


def main():
    print("Selecione o arquivo XML")
    file_name_xml = recebe_nome_do_arquivo()
    file_name = file_name_xml.replace(".XML", "").replace(".xml", "")
    corrige_caracteres_especiais(file_name_xml)
    tree_xml = ElementTree.parse(file_name_xml)
    root = tree_xml.getroot()

    # Cria um arquivo excel e adiciona uma aba para cada informação
    workbook = xlsxwriter.Workbook(f'{file_name}.xlsx')
    for child in root:
        worksheet = workbook.add_worksheet(child.tag)

        # Separa uma linha para cada instância, deixando a primeira para o cabeçalho
        for row, gchild in enumerate(child):

            # Imprime os atributos de cada instância
            for col, attribute in enumerate(gchild.attrib):

                # Na primeira linha imprime o cabeçalho
                if row == 0:
                    worksheet.write(row, col, attribute)

                worksheet.write(row + 1, col, gchild.attrib[attribute])

    workbook.close()
    print("Arquivo gerado com sucesso!")


if __name__ == "__main__":
    main()
