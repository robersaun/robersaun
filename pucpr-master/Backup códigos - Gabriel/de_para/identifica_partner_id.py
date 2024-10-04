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
    file_data = file_data.replace('P&D', 'P&amp;D')

    # Sobrescreve o arquivo
    with open(file_name, 'w') as file:
        file.write(file_data)


def encontra_disciplinas(tree):
    root = tree.getroot()

    disciplinas = []
    for child in root.iter("subject"):
        nome = child.get("name")
        short = child.get("short")
        partner_id = child.get("partner_id")
        disciplinas.append([nome, short, partner_id])

    return disciplinas


def encontra_turmas(tree):
    root = tree.getroot()

    turmas = []
    for child in root.iter("class"):
        nome = child.get("name")
        short = child.get("short")
        partner_id = child.get("partner_id")
        turmas.append([nome, short, partner_id])

    return turmas


def encontra_professores(tree):
    root = tree.getroot()

    professores = []
    for child in root.iter("teacher"):
        nome = child.get("name")
        partner_id = child.get("customfield1")
        professores.append([nome, partner_id])

    return professores


def main():
    print("Selecione o arquivo XML")
    file_name_xml = recebe_nome_do_arquivo()
    file_name = file_name_xml.replace(".XML", "").replace(".xml", "")
    corrige_caracteres_especiais(file_name_xml)
    tree_xml = ElementTree.parse(file_name_xml)
    disciplinas_xml = encontra_disciplinas(tree_xml)
    turmas_xml = encontra_turmas(tree_xml)
    professores_xml = encontra_professores(tree_xml)

    # Cria um arquivo excel e adiciona três abas
    workbook = xlsxwriter.Workbook(f'{file_name}.xlsx')
    worksheet_disciplinas = workbook.add_worksheet("Disciplinas")
    worksheet_turmas = workbook.add_worksheet("Turmas")
    worksheet_professores = workbook.add_worksheet("Professores")

    # ## Gravação de disciplinas ## #

    # A primeira linha do arquivo fica reservada para o cabeçalho
    row = 1
    col = 0

    # Itera através dos dados escrevendo linha por linha
    for nome, short, partner_id in disciplinas_xml:
        worksheet_disciplinas.write(row, col, nome)
        worksheet_disciplinas.write(row, col + 1, short)
        worksheet_disciplinas.write(row, col + 2, partner_id)
        row += 1

    # Formata a tabela (cabeçalho e largura das colunas)
    worksheet_disciplinas.add_table(0, 0, row - 1, 2, {'style': 'Table Style Light 1', 'columns': [
        {'header': 'Nome'},
        {'header': 'Abreviação'},
        {'header': 'Partner ID'},
    ]})
    worksheet_disciplinas.set_column('A:D', 50)

    # ## Gravação de turmas ## #

    # A primeira linha do arquivo fica reservada para o cabeçalho
    row = 1
    col = 0

    # Itera através dos dados escrevendo linha por linha
    for nome, short, partner_id in turmas_xml:
        worksheet_turmas.write(row, col, nome)
        worksheet_turmas.write(row, col + 1, short)
        worksheet_turmas.write(row, col + 2, partner_id)
        row += 1

    # Formata a tabela (cabeçalho e largura das colunas)
    worksheet_turmas.add_table(0, 0, row - 1, 2, {'style': 'Table Style Light 1', 'columns': [
        {'header': 'Nome'},
        {'header': 'Abreviação'},
        {'header': 'Partner ID'},
    ]})
    worksheet_turmas.set_column('A:C', 50)

    # ## Gravação de professores ## #

    # A primeira linha do arquivo fica reservada para o cabeçalho
    row = 1
    col = 0

    # Itera através dos dados escrevendo linha por linha
    for nome, partner_id in professores_xml:
        worksheet_professores.write(row, col, nome)
        worksheet_professores.write(row, col + 1, partner_id)
        row += 1

    # Formata a tabela (cabeçalho e largura das colunas)
    worksheet_professores.add_table(0, 0, row - 1, 1, {'style': 'Table Style Light 1', 'columns': [
        {'header': 'Nome'},
        {'header': 'Partner ID'},
    ]})
    worksheet_professores.set_column('A:B', 50)

    workbook.close()
    print("Relatório gerado com sucesso!")


if __name__ == "__main__":
    main()
