'''
Recebe um XML do Prime e corrige os caracteres especiais para realizar a importação para o ASC

Desenvolvido por Vinicius Tozo
Última atualização: 18/11/2020
'''

import xml.etree.ElementTree as ElementTree
from tkinter import Tk
from tkinter.filedialog import askopenfilename


def get_file_name():
    Tk().withdraw()
    file_name = askopenfilename(filetypes=[('XML', '.xml')])
    return file_name


def fix_special_characters(file_name):
    # Read in the file
    with open(file_name, 'r', encoding='windows-1252') as file:
        file_data = file.read()

    # Replace the target string
    file_data = file_data.replace(' & ', ' &amp; ')

    # Write the file out again
    with open(file_name, 'w') as file:
        file.write(file_data)


def set_short_as_name(tree):
    root = tree.getroot()
    # Iterate through all subject elements
    element_filter = "subject"
    for child in root.iter(element_filter):
        # Copy contents from name to short
        child.set("short", child.get("name"))
    return tree


def main():
    print("Selecione o arquivo para ser corrigido")
    file_name = get_file_name()

    fix_special_characters(file_name)
    tree = ElementTree.parse(file_name)
    tree = set_short_as_name(tree)

    # Save changes to the selected file
    tree.write(file_name, encoding='windows-1252')

    print("Alterações concluídas com sucesso!")
    input("Pressione Enter para fechar esta janela...")


if __name__ == "__main__":
    main()
