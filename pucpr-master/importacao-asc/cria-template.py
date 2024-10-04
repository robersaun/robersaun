'''
Recebe um XML do Prime e um XML do ASC  e une as informações das disciplinas e turmas em uma tabela Excel

Desenvolvido por Vinicius Tozo
Última atualizalção: 12/07/2021

'''

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


def encontra_disciplinas(tree):
    root = tree.getroot()

    disciplinas = []
    for child in root.iter("subject"):
        nome = child.get("name")
        short = child.get("short")
        partner_id = child.get("partner_id")
        validador = ";".join(nome.split(";")[0:-1])
        disciplinas.append([nome, short, validador, partner_id])

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


def main():
    print("Selecione o arquivo XML do aSc Timetables")
    file_name_asc = recebe_nome_do_arquivo()
    corrige_caracteres_especiais(file_name_asc)
    tree_asc = ElementTree.parse(file_name_asc)
    disciplinas_asc = encontra_disciplinas(tree_asc)
    turmas_asc = encontra_turmas(tree_asc)

    print("Selecione o arquivo XML do Prime")
    file_name_prime = recebe_nome_do_arquivo()
    corrige_caracteres_especiais(file_name_prime)
    tree_prime = ElementTree.parse(file_name_prime)
    disciplinas_prime = encontra_disciplinas(tree_prime)
    turmas_prime = encontra_turmas(tree_prime)

    # Cria um arquivo excel e adiciona quatro abas
    workbook = xlsxwriter.Workbook('Template de Importação aSc.xlsx')
    worksheet_disciplinas_asc = workbook.add_worksheet("Disciplinas aSc")
    worksheet_disciplinas_prime = workbook.add_worksheet("Disciplinas Prime")
    worksheet_turmas_asc = workbook.add_worksheet("Turmas aSc")
    worksheet_turmas_prime = workbook.add_worksheet("Turmas Prime")
    worksheet_importacao_disciplinas = workbook.add_worksheet("Importação Disciplinas")
    worksheet_importacao_turmas = workbook.add_worksheet("Importação Turmas")

    # ## Gravação de disciplinas aSc ## #

    # A primeira linha do arquivo fica reservada para o cabeçalho
    row = 1
    col = 0

    # Itera através dos dados escrevendo linha por linha
    for nome, short, validador, partner_id in disciplinas_asc:
        worksheet_disciplinas_asc.write(row, col, validador)
        worksheet_disciplinas_asc.write(row, col + 1, nome)
        worksheet_disciplinas_asc.write(row, col + 2, short)
        worksheet_disciplinas_asc.write_formula(row, col + 3, f"=VLOOKUP(A{row + 1}, 'Disciplinas Prime'!A:B, 2, 0)")
        worksheet_disciplinas_asc.write(row, col + 4, partner_id)
        row += 1

    # Formata a tabela (cabeçalho e largura das colunas)
    worksheet_disciplinas_asc.add_table(0, 0, row - 1, 4, {'style': 'Table Style Light 1', 'columns': [
        {'header': 'Validador'},
        {'header': 'Nome'},
        {'header': 'Abreviação'},
        {'header': 'PROCV Prime'},
        {'header': 'Partner ID'},
    ]})
    worksheet_disciplinas_asc.set_column('A:E', 50)

    # ## Gravação de disciplinas Prime ## #

    # A primeira linha do arquivo fica reservada para o cabeçalho
    row = 1
    col = 0

    # Itera através dos dados escrevendo linha por linha
    for nome, short, validador, partner_id in disciplinas_prime:
        worksheet_disciplinas_prime.write(row, col, validador)
        worksheet_disciplinas_prime.write(row, col + 1, nome)
        worksheet_disciplinas_prime.write(row, col + 2, short)
        worksheet_disciplinas_prime.write(row, col + 3, partner_id)
        row += 1

    # Formata a tabela (cabeçalho e largura das colunas)
    worksheet_disciplinas_prime.add_table(0, 0, row - 1, 3, {'style': 'Table Style Light 1', 'columns': [
        {'header': 'Validador'},
        {'header': 'Nome'},
        {'header': 'Abreviação'},
        {'header': 'Partner ID'},
    ]})
    worksheet_disciplinas_prime.set_column('A:D', 50)

    # ## Gravação de turmas aSc ## #

    # A primeira linha do arquivo fica reservada para o cabeçalho
    row = 1
    col = 0

    # Itera através dos dados escrevendo linha por linha
    for nome, short, partner_id in turmas_asc:
        worksheet_turmas_asc.write(row, col, nome)
        worksheet_turmas_asc.write(row, col + 1, short)
        worksheet_turmas_asc.write_formula(row, col + 2, f"=VLOOKUP(A{row + 1}, 'Turmas Prime'!A:B, 2, 0)")
        worksheet_turmas_asc.write(row, col + 3, partner_id)
        row += 1

    # Formata a tabela (cabeçalho e largura das colunas)
    worksheet_turmas_asc.add_table(0, 0, row - 1, 3, {'style': 'Table Style Light 1', 'columns': [
        {'header': 'Nome'},
        {'header': 'Abreviação'},
        {'header': 'PROCV Prime'},
        {'header': 'Partner ID'},
    ]})
    worksheet_turmas_asc.set_column('A:D', 50)

    # ## Gravação de turmas Prime ## #

    # A primeira linha do arquivo fica reservada para o cabeçalho
    row = 1
    col = 0

    # Itera através dos dados escrevendo linha por linha
    for nome, short, partner_id in turmas_prime:
        worksheet_turmas_prime.write(row, col, nome)
        worksheet_turmas_prime.write(row, col + 1, short)
        worksheet_turmas_prime.write(row, col + 2, partner_id)
        row += 1

    # Formata a tabela (cabeçalho e largura das colunas)
    worksheet_turmas_prime.add_table(0, 0, row - 1, 2, {'style': 'Table Style Light 1', 'columns': [
        {'header': 'Nome'},
        {'header': 'Abreviação'},
        {'header': 'Partner ID'},
    ]})
    worksheet_turmas_prime.set_column('A:D', 50)

    # ## Gravação de importações de Disciplinas ## #

    bold_format = workbook.add_format({'bold': True})

    # As primeiras duas linhas do arquivo ficam reservadas para o cabeçalho
    row = 2
    col = 0

    worksheet_importacao_disciplinas.write(1, 0, "Nome", bold_format)
    worksheet_importacao_disciplinas.write(1, 1, "Abreviação", bold_format)
    worksheet_importacao_disciplinas.write(1, 2, "Nome", bold_format)
    worksheet_importacao_disciplinas.write(1, 3, "Abreviação", bold_format)

    # Itera através dos dados escrevendo linha por linha
    contagem_de_disciplinas = len(disciplinas_asc)
    for i in range(0, contagem_de_disciplinas):
        worksheet_importacao_disciplinas.write(row, col, f"='Disciplinas aSc'!B{row}")
        worksheet_importacao_disciplinas.write_formula(row, col + 1, f"='Disciplinas aSc'!D{row}")
        worksheet_importacao_disciplinas.write_formula(row, col + 2, f"='Disciplinas aSc'!D{row}")
        worksheet_importacao_disciplinas.write_formula(row, col + 3, f"='Disciplinas aSc'!D{row}")
        row += 1

    # Formata a tabela (cabeçalho e largura das colunas)
    worksheet_importacao_disciplinas.add_table(0, 0, row - 1, 3, {'style': 'Table Style Light 1', 'columns': [
        {'header': 'Importação 1 - Nome aSc'},
        {'header': 'Importação 1 - Abreviação Prime'},
        {'header': 'Importação 2 - Nome Prime'},
        {'header': 'Importação 2 - Abreviação Prime'},
    ]})
    worksheet_importacao_disciplinas.set_column('A:D', 50)

    # ## Gravação de importações de Turmas ## #

    # As primeiras duas linhas do arquivo ficam reservadas para o cabeçalho
    row = 2
    col = 0

    worksheet_importacao_turmas.write(1, 0, "Nome", bold_format)
    worksheet_importacao_turmas.write(1, 1, "Abreviação", bold_format)
    worksheet_importacao_turmas.write(1, 2, "Nome", bold_format)
    worksheet_importacao_turmas.write(1, 3, "Abreviação", bold_format)

    # Itera através dos dados escrevendo linha por linha
    contagem_de_turmas = len(turmas_asc)
    for i in range(0, contagem_de_turmas):
        worksheet_importacao_turmas.write(row, col, f"='Turmas aSc'!A{row}")
        worksheet_importacao_turmas.write_formula(row, col + 1, f"='Turmas aSc'!C{row}")
        worksheet_importacao_turmas.write_formula(row, col + 2, f"='Turmas aSc'!C{row}")
        worksheet_importacao_turmas.write_formula(row, col + 3, f"='Turmas aSc'!C{row}")
        row += 1

    # Formata a tabela (cabeçalho e largura das colunas)
    worksheet_importacao_turmas.add_table(0, 0, row - 1, 3, {'style': 'Table Style Light 1', 'columns': [
        {'header': 'Importação 1 - Nome aSc'},
        {'header': 'Importação 1 - Abreviação Prime'},
        {'header': 'Importação 2 - Nome Prime'},
        {'header': 'Importação 2 - Abreviação Prime'},
    ]})
    worksheet_importacao_turmas.set_column('A:D', 50)

    workbook.close()
    print("Template gerado com sucesso!")


if __name__ == "__main__":
    main()
