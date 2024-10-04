'''
Script recebe um Excel com informações dos docentes e gera um arquivo com os dados dos docentes formatados (cpf e data de nascimento) e com os seus vínculos

Desenvolvido por Matheus Rosa e Vinicius Tozo
Última atualização: 19/08/2021
'''

import xml.etree.ElementTree as ElementTree
from tkinter import Tk
from tkinter.filedialog import askopenfilename


def fix_subjects(root):
    aula_perm = ['9999;99;PERMANÊNCIA;PER', '9999;99;PERMANÊNCIA EXTERNA;PEREXT']
    for subject in root.iter("subject"):

        # retira o * do name se houver
        if len(subject.get("name").split(";")) >= 3:
            parts = subject.get("name").split(";")
            if parts[2][-1:] == "*":
                parts[2] = parts[2][:-3]
                new_name = ";".join(parts)
                subject.set("name", new_name)
    # excluindo permanencias
    for subject in root.find("subjects").findall("subject"):
        if subject.get("name") in aula_perm:
            print('Aviso: a disciplina', subject.get("name"), 'foi EXCLUIDA')
            root.find("subjects").remove(subject)


def converte_xml(file_name_xml):
    tree_xml = ElementTree.parse(file_name_xml)
    root = tree_xml.getroot()

    fix_subjects(root)

    return tree_xml


def main():
    # Lê o arquivo XML no formato Prime
    print("Selecione o arquivo XML")
    Tk().withdraw()
    arquivo_xml = askopenfilename(filetypes=[("XML", ".xml")])
    print(f"\tArquivo selecionado: {arquivo_xml}")

    # Converte o formato do XML do formato Prime para o formato Ponto
    xml_convertido = converte_xml(arquivo_xml)

    # Salva o arquivo formatado
    xml_convertido.write("arquivo_prime_gerado.xml", encoding="windows-1252")

    input("Arquivo gerado com sucesso!")


if __name__ == "__main__":
    main()
