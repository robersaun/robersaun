'''
Recebe um Excel e converte para um arquivo XML ASC

Desenvolvido por Vinicius Tozo
Última atualização: 16/07/2021
'''

import xml.etree.ElementTree as ElementTree
import pandas
from xml.dom import minidom
from tkinter import Tk
from tkinter.filedialog import askopenfilename


def recebe_nome_do_arquivo():
    Tk().withdraw()
    file_name = askopenfilename(filetypes=[('xlsx', '.xlsx')])
    return file_name


def main():
    print("Selecione o arquivo excel")
    file_name_xlsx = recebe_nome_do_arquivo()
    file_name = file_name_xlsx.replace(".XLSX", ".xml").replace(".xlsx", ".xml")

    root = ElementTree.Element(
        "timetable", ascttversion="2020.7.5", importtype="database",
        options="export:idprefix:%CHRID,import:idprefix:%TEMPID,groupstype1,decimalseparatordot,"
                "lessonsincludeclasseswithoutstudents,handlestudentsafterlessons",
        defaultexport="1", displayname="aSc Timetables 2012 XML", displaycountries="")

    excel = pandas.ExcelFile(file_name_xlsx)

    for sheet in excel.sheet_names:  # Para cada aba no arquivo excel
        child_name = sheet[:-1]

        if child_name == "classe":
            child_name = "class"

        columns = excel.parse(sheet).columns
        converters = {column: str for column in columns}

        dataframe = excel.parse(sheet, converters=converters).fillna("")

        columns = ",".join(dataframe.columns.tolist())

        element = ElementTree.SubElement(root, sheet, options="canadd,export:silent", columns=columns)

        for index, row in dataframe.iterrows():
            child = ElementTree.SubElement(element, child_name)
            for column in columns.split(","):
                child.set(column, str(row[column]))

    # Salva o arquivo formatado
    xml_string = minidom.parseString(ElementTree.tostring(root, encoding="windows-1252")).toprettyxml(indent="   ")
    with open(file_name, "w", encoding="windows-1252") as f:
        f.write(xml_string)


if __name__ == '__main__':
    main()
