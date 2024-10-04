'''
Recebe um Excel e um XML fazendo a adição do partner id em cada disciplina corretamente

Desenvolvido por Vinicius Tozo
Última atualização: 30/06/2021
'''

from tkinter import Tk
from tkinter.filedialog import askopenfilename
from xml.etree import ElementTree

import pandas


def cria_dicionario_disciplina_partner_id(arquivo):
    # Lê o arquivo excel que contêm os dados
    dataframe = pandas.read_excel(arquivo, engine="openpyxl")

    # Cria um dicionário com os dados do excel
    dicionario = {}
    for index, row in dataframe.iterrows():
        name = str(row["Nome"]).strip()
        partner_id = str(row["Id"]).strip()

        # Se não contém só número
        if not partner_id.isdigit():
            continue

        dicionario[name] = partner_id

    return dicionario


def insere_partner_id(arquivo_xml, dicionario_de_disciplinas):
    tree_xml = ElementTree.parse(arquivo_xml)
    root = tree_xml.getroot()

    for subject in root.iter("subject"):
        if subject.get("partner_id") != "":
            continue
        name = subject.get("name")
        subject.set("partner_id", dicionario_de_disciplinas.get(name, ""))

    return tree_xml


def main():
    Tk().withdraw()
    arquivo_xlsx = askopenfilename(filetypes=[("xlsx", ".xlsx")])
    dicionario_de_disciplinas = cria_dicionario_disciplina_partner_id(arquivo_xlsx)

    Tk().withdraw()
    arquivo_xml = askopenfilename(filetypes=[("xml", ".xml")])
    xml_corrigido = insere_partner_id(arquivo_xml, dicionario_de_disciplinas)
    xml_corrigido.write("arquivo_gerado.xml", encoding="windows-1252")


if __name__ == '__main__':
    main()
