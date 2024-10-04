"""Comparador de Matrizes do Prime com Resolução em Lote

Script para comparar todas as matrizes baixada do Prime com a respectiva resolução.
Recebe como entrada a pasta raiz contendo todos os arquivos de resolução e de matrizes em formato XLSX,
e gera um relatório de diferenças no formato Excel dentro de cada pasta.
As resoluções originalmente vêm em formato Word, e para converter para Excel deve-se copiar todos os dados e colar em
uma nova planilha, cuidando para que cada disciplina fique em uma única linha.
As matrizes do Prime originalmente vêm em formato XLS, para converter em XLSX deve-se abrí-las no Excel e 'Salvar como'.

É utilizado pelo analista de TI no processo de sanitização de matrizes.


Desenvolvido por Vinicius Tozo
Última atualização: 13/08/2021
"""
import os
import pandas
import xlrd
import re
from tkinter import filedialog
from tkinter import Tk


def encontra_matrizes_a_fazer(root):
    contagem = 0

    # Para cada diretório dentro do diretório raiz
    for dirpath, dirnames, filenames in os.walk(root):

        # Se o diretório possuir outros diretórios dentro dele, ignora
        if dirnames:
            continue

        # Se não, conta como uma matriz e continua a contagem de arquivos

        possui_relatorio_erros = False
        possui_resolucao = False
        possui_matriz = False
        caminho_matriz = ""
        caminho_resolucao = ""

        # Para cada arquivo dentro do diretório, verifica se a matriz possui os arquivos necessários
        for file in filenames:
            if file.startswith("Grade_curricular_") and file.endswith(".xlsx"):
                possui_matriz = True
                caminho_matriz = file
            elif file == "Novo(a) Planilha do Microsoft Excel.xlsx":
                possui_resolucao = True
                caminho_resolucao = file
            elif file == "Relatório de Erros.xlsx":
                possui_relatorio_erros = True

        if possui_matriz and possui_resolucao and not possui_relatorio_erros:
            # Se possuir os arquivos, tenta comparar as matrizes
            try:
                compara_matrizes(root, dirpath + "/" + caminho_matriz, dirpath + "/" + caminho_resolucao)
                contagem += 1
            except Exception as e:
                print("Ocorreu um erro: ", str(e))
                print(f"na matriz {dirpath}")
                print()

    return contagem


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


def compara_matrizes(root, arquivo_excel, arquivo_word):
    column_names = ["Coluna(s) com erro", "Ordem", "Ciclo", "Disciplina", "Requisitos", "Carga horária teórica",
                    "Carga horária prática", "Créditos", "Carga Horária oficial", "Carga Horária relógio oficial"]
    nome_arquivo_saida = \
        arquivo_word.replace("/Novo(a) Planilha do Microsoft Excel.xlsx", "") + "/" + "Relatório de Erros.xlsx"

    workbook_excel = xlrd.open_workbook(arquivo_excel)
    sheet_excel = workbook_excel.sheet_by_name("Ciclos")

    workbook_word = xlrd.open_workbook(arquivo_word)
    sheet_word = workbook_word.sheet_by_name("Plan1")

    disciplinas_excel = []
    disciplinas_word = []

    # Dados Word
    ciclo = 0
    ordem_word = []

    for linha in range(0, sheet_word.nrows):
        if type(sheet_word.cell_value(linha, 0)) is str:
            if "PERÍODO" in str.upper(sheet_word.cell_value(linha, 0)):
                ciclo = int(
                    sheet_word.cell_value(linha, 0).replace("°", "|").replace("º", "|").replace(".", "").split("|")[0])
            continue
        ordem_disciplina = int(sheet_word.cell_value(linha, 0))
        nome_disciplina = sheet_word.cell_value(linha, 1).strip().replace("–", "-")
        requisito_disciplina = sheet_word.cell_value(linha, 2) \
            .replace(" ", "").replace("\xa0", "").replace(";", ",").strip("0")
        aula_teorica = int(sheet_word.cell_value(linha, 3))
        aula_pratica = int(sheet_word.cell_value(linha, 4))
        creditos = int(sheet_word.cell_value(linha, 5))
        ch_oficial = int(sheet_word.cell_value(linha, 6))
        ch_relogio_oficial = int(sheet_word.cell_value(linha, 7))

        disciplinas_word.append([
            ciclo, nome_disciplina, requisito_disciplina, aula_teorica,
            aula_pratica, creditos, ch_oficial, ch_relogio_oficial])
        ordem_word.append(ordem_disciplina)

    if not disciplinas_word:
        raise Exception("Erro na leitura da resolução")

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

        requisito_disciplina = tratar_requisitos(sheet_excel.cell_value(linha, 17), disciplinas_word)

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

        disciplinas_excel.append(
            [ciclo, nome_disciplina, requisito_disciplina, aula_teorica, aula_pratica, creditos, ch_oficial,
             ch_relogio_oficial])

    if not disciplinas_excel:
        raise Exception("Erro na leitura da matriz")

    consistente = 0
    inconsistente = 0
    resultado = []

    for index, word in enumerate(disciplinas_word):
        if "Eletiva" in word:
            continue
        elif word in disciplinas_excel:
            consistente += 1
        else:
            erro = []
            existe_no_excel = any(word[1] in sublist for sublist in disciplinas_excel)
            if existe_no_excel:
                indice = [x for x in disciplinas_excel if word[1] in x][0]
                excel = disciplinas_excel[disciplinas_excel.index(indice)]
                if word[0] != excel[0]:
                    erro.append("ciclo")
                if word[2] != excel[2] and excel[2].__contains__("???"):
                    erro.append("requisitos (pode ser erro no nome de uma outra disciplina)")
                if word[2] != excel[2] and not excel[2].__contains__("???"):
                    erro.append("requisitos")
                if word[3] != excel[3]:
                    erro.append("carga horária teórica")
                if word[4] != excel[4]:
                    erro.append("carga horária prática")
                if word[5] != excel[5]:
                    erro.append("créditos")
                if word[6] != excel[6]:
                    erro.append("carga horária oficial")
                if word[7] != excel[7]:
                    erro.append("carga horária relógio oficial")
            else:
                erro.append("nome da disciplina (verificar todos os campos)")
            separador_nome = ", "
            erros = separador_nome.join(erro)
            word.insert(0, ordem_word[index])
            word.insert(0, erros)
            resultado.append(word)
            inconsistente += 1

    print(nome_arquivo_saida.replace(root, ""))
    df_saida = pandas.DataFrame(columns=column_names, data=resultado)
    df_saida.to_excel(nome_arquivo_saida, sheet_name='Relatório', index=False)

    print("Valores consistentes: " + str(consistente))
    print("Valores inconsistentes: " + str(inconsistente) + "\n")


if __name__ == '__main__':
    Tk().withdraw()
    pasta_raiz = filedialog.askdirectory()
    print()
    print()
    print("Relatórios feitos: ", encontra_matrizes_a_fazer(pasta_raiz))
