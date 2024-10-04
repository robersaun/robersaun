"""Comparador de Matrizes do Prime com Resolução

Script para comparar uma única matriz baixada do Prime com uma resolução.
Recebe como entrada ambos os arquivos em formato Excel, e gera um relatório de diferenças também em formato Excel.
As resoluções originalmente vêm em formato Word, e para converter para Excel deve-se copiar todos os dados e colar em
uma nova planilha, cuidando para que cada disciplina fique em uma única linha.
As matrizes do Prime originalmente vêm em formato XLS, para converter em XLSX deve-se abrí-las no Excel e 'Salvar como'.

Pode ser utilizado pelo analista de TI para auxiliar no processo de sanitização de matrizes.

Desenvolvido por Vinicius Tozo
Última atualização: 13/09/2021
"""
import re
from tkinter import Tk
from tkinter.filedialog import askopenfilename

import pandas
import xlrd


def tratar_requisitos(req, disciplinas_word):
    resultados = []
    requisitos = re.split("Pré-R|Co-Re|Requisito de Posição: |Requisito Espe", req)

    for requisito in requisitos:
        requisito = requisito.replace(",", "").strip()
        if "equisito: " in requisito:  # Pré-requisito
            disciplina = requisito.strip().split(" - ")[1:]
            separador = " – "
            disciplina = separador.join(disciplina).strip().replace("–", "-")
            existe = any(disciplina in sublist for sublist in disciplinas_word)

            if existe:
                x = [x for x in disciplinas_word if disciplina in x][0]
                ordem = disciplinas_word.index(x) + 1
            else:
                ordem = "???"
            resultados.append(str(ordem) + "PR")
        elif "quisito Direcional: " in requisito:  # Co-requisito Direcional
            disciplina = requisito.strip().split(" - ")[1:]
            separador = " – "
            disciplina = separador.join(disciplina).strip().replace("–", "-")
            existe = any(disciplina in sublist for sublist in disciplinas_word)

            if existe:
                x = [x for x in disciplinas_word if disciplina in x][0]
                ordem = disciplinas_word.index(x) + 1
            else:
                ordem = "???"
            resultados.append(str(ordem) + "CD")
        elif "quisito: " in requisito:  # Co-requisito
            disciplina = requisito.strip().split(" - ")[1:]
            separador = " – "
            disciplina = separador.join(disciplina).strip().replace("–", "-")
            existe = any(disciplina in sublist for sublist in disciplinas_word)

            if existe:
                x = [x for x in disciplinas_word if disciplina in x][0]
                ordem = disciplinas_word.index(x) + 1
            else:
                ordem = "???"
            resultados.append(str(ordem) + "CR")
        elif "cial: " in requisito:  # Requisito especial
            disciplina = requisito.strip().split(" - ")[1:]
            separador = " – "
            disciplina = separador.join(disciplina).strip().replace("–", "-")
            existe = any(disciplina in sublist for sublist in disciplinas_word)

            if existe:
                x = [x for x in disciplinas_word if disciplina in x][0]
                ordem = disciplinas_word.index(x) + 1
            else:
                ordem = "???"
            resultados.append(str(ordem) + "RE")
        elif "Percentual Mínimo" in requisito:
            percentual = requisito.split("= ")[-1]
            if percentual != "0":
                resultados.append("RP" + str(percentual) + "%")
        elif "Número Mínimo de Créditos Cursados" in requisito:
            n = requisito.split("= ")[-1]
            if n != "0":
                resultados.append("RP" + str(n) + "ct")
        elif "Total Mínimo de Horas Relógio Cursadas" in requisito:
            n = requisito.split("= ")[-1]
            if n != "0":
                resultados.append("RP" + str(n) + "HR")

    separador = ","
    return separador.join(resultados)


def compara_matrizes(disciplinas_prime, disciplinas_resolucao, root, ordem_word):
    column_names = ["Coluna(s) com erro", "Ordem", "Ciclo", "Disciplina", "Requisitos", "Carga horária teórica",
                    "Carga horária prática", "Créditos", "Carga Horária oficial", "Carga Horária relógio oficial"]

    nome_arquivo_saida = root + "Relatório de Erros.xlsx"

    consistente = 0
    inconsistente = 0
    resultado = []

    print("\n----Grade Prime----")
    for i in disciplinas_prime:
        print(i)

    print("\n----Resolução----")
    for i in disciplinas_resolucao:
        print(i)

    print("\n----Relatório (dados da resolução)----")

    for index, disciplina_resolucao in enumerate(disciplinas_resolucao):
        if "Eletiva" in disciplina_resolucao:
            continue
        elif disciplina_resolucao in disciplinas_prime:
            consistente += 1
        else:
            erro = []
            existe_no_excel = any(disciplina_resolucao[1] in sublist for sublist in disciplinas_prime)
            if existe_no_excel:
                indice = [x for x in disciplinas_prime if disciplina_resolucao[1] in x][0]
                excel = disciplinas_prime[disciplinas_prime.index(indice)]
                if disciplina_resolucao[0] != excel[0]:
                    erro.append("ciclo")
                if disciplina_resolucao[2] != excel[2] and excel[2].__contains__("???"):
                    erro.append("requisitos (pode ser erro no nome de uma outra disciplina)")
                if disciplina_resolucao[2] != excel[2] and not excel[2].__contains__("???"):
                    erro.append("requisitos")
                if disciplina_resolucao[3] != excel[3]:
                    erro.append("carga horária teórica")
                if disciplina_resolucao[4] != excel[4]:
                    erro.append("carga horária prática")
                if disciplina_resolucao[5] != excel[5]:
                    erro.append("créditos")
                if disciplina_resolucao[6] != excel[6]:
                    erro.append("carga horária oficial")
                if disciplina_resolucao[7] != excel[7]:
                    erro.append("carga horária relógio oficial")
            else:
                erro.append("nome da disciplina (verificar todos os campos)")
            separador_nome = ", "
            erros = separador_nome.join(erro)
            disciplina_resolucao.insert(0, ordem_word[index])
            disciplina_resolucao.insert(0, erros)
            resultado.append(disciplina_resolucao)
            print(disciplina_resolucao)
            inconsistente += 1

    df_saida = pandas.DataFrame(columns=column_names, data=resultado)
    df_saida.to_excel(nome_arquivo_saida, sheet_name='Relatório', index=False)

    print()
    print("Valores consistentes: " + str(consistente))
    print("Valores inconsistentes: " + str(inconsistente) + "\n")
    print()
    print(root)


def ler_resolucao(caminho_resolucao):
    workbook_word = xlrd.open_workbook(caminho_resolucao)
    sheet_word = workbook_word.sheet_by_name("Plan1")
    disciplinas_resolucao = []

    # Dados Word
    ciclo = 0
    ordem_word = []

    for linha in range(0, sheet_word.nrows):
        if "PERÍODO" in str.upper(sheet_word.cell_value(linha, 0)):
            ciclo = int(''.join(filter(str.isdigit, sheet_word.cell_value(linha, 0))))
        elif "PERÍODO" in str.upper(str(sheet_word.cell_value(linha, 1))):
            ciclo = int(''.join(filter(str.isdigit, sheet_word.cell_value(linha, 1))))
        try:
            int(sheet_word.cell_value(linha, 0).__str__().strip())
        except ValueError:
            continue

        ordem_disciplina = int(sheet_word.cell_value(linha, 0).__str__().strip())
        nome_disciplina = sheet_word.cell_value(linha, 1).strip().replace("–", "-")
        nome_disciplina = nome_disciplina.replace(u'\xa0', u' ')  # Caractere especial (&nbsp;)
        requisito_disciplina = sheet_word.cell_value(linha, 2) \
            .replace(" ", "").replace("\xa0", "").replace("\u202f", "").replace(";", ",").strip("0")
        aula_teorica = int(sheet_word.cell_value(linha, 3))
        aula_pratica = int(sheet_word.cell_value(linha, 4))
        creditos = int(sheet_word.cell_value(linha, 5))
        ch_oficial = int(sheet_word.cell_value(linha, 6))
        ch_relogio_oficial = int(sheet_word.cell_value(linha, 7))

        disciplinas_resolucao.append([
            ciclo, nome_disciplina, requisito_disciplina, aula_teorica,
            aula_pratica, creditos, ch_oficial, ch_relogio_oficial])
        ordem_word.append(ordem_disciplina)

    if not disciplinas_resolucao:
        raise Exception("Erro na leitura da resolução")

    return disciplinas_resolucao, ordem_word


def ler_grade(caminho_prime, disciplinas_resolucao):
    workbook_excel = xlrd.open_workbook(caminho_prime)
    sheet_excel = workbook_excel.sheet_by_name("Ciclos")

    disciplinas_prime = []

    # Dados Excel
    ciclo = 0

    for linha in range(0, sheet_excel.nrows):
        if sheet_excel.cell_value(linha, 1) == "" or \
                sheet_excel.cell_value(linha, 1) == "Total Carga Horária (mínimo):" or \
                sheet_excel.cell_value(linha, 1) == "Disciplina" or \
                sheet_excel.cell_value(linha, 1) == "Total":
            continue
        if sheet_excel.cell_value(linha, 1) == "Ciclo:":
            ciclo = int(sheet_excel.cell_value(linha, 2))
            continue

        nome_disciplina = sheet_excel.cell_value(linha, 1).strip().split(" - ")[1:]
        separador_nome = " – "
        nome_disciplina = separador_nome.join(nome_disciplina).strip().replace("–", "-")

        requisito_disciplina = tratar_requisitos(sheet_excel.cell_value(linha, 17), disciplinas_resolucao)

        aula_teorica = sheet_excel.cell_value(linha, 6)
        aula_pratica = sheet_excel.cell_value(linha, 7)
        creditos = sheet_excel.cell_value(linha, 15)
        ch_oficial = sheet_excel.cell_value(linha, 8)
        ch_relogio_oficial = sheet_excel.cell_value(linha, 12)

        if aula_teorica == "":
            aula_teorica = 0
        elif int(aula_teorica) % 20 == 0:
            aula_teorica = int(aula_teorica) / 20
        else:
            aula_teorica = int(aula_teorica) / 18

        if aula_pratica == "":
            aula_pratica = 0
        elif int(aula_pratica) % 20 == 0:
            aula_pratica = int(aula_pratica) / 20
        else:
            aula_pratica = int(aula_pratica) / 18

        if creditos == "":
            creditos = 0
        else:
            creditos = int(creditos)

        if ch_oficial == "":
            ch_oficial = 0
        else:
            ch_oficial = int(ch_oficial)

        if ch_relogio_oficial == "":
            ch_relogio_oficial = 0
        else:
            ch_relogio_oficial = int(ch_relogio_oficial)

        disciplinas_prime.append(
            [ciclo, nome_disciplina, requisito_disciplina, aula_teorica, aula_pratica, creditos, ch_oficial,
             ch_relogio_oficial])

    if not disciplinas_prime:
        raise Exception("Erro na leitura da matriz")

    return disciplinas_prime


def main():
    Tk().withdraw()
    caminho_resolucao = askopenfilename(
        filetypes=[('Arquivo excel', '.xlsx')],
        title='Selecione o arquivo da resolução')
    resolucao, ordem_word = ler_resolucao(caminho_resolucao)

    caminho_prime = askopenfilename(
        filetypes=[('Arquivo excel', '.xlsx')],
        title='Selecione o arquivo extraído do Prime')
    grade_prime = ler_grade(caminho_prime, resolucao)

    root = "/".join(caminho_resolucao.split("/")[0:-1]) + "/"
    compara_matrizes(grade_prime, resolucao, root, ordem_word)


if __name__ == '__main__':
    main()
