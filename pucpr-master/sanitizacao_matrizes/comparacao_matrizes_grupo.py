'''
Script realiza a comparação entre matrizes, verificando se há incosistencias  e gerando arquvio Excel no final.
Para realizar isso, é necessário que o script receba 1 excel com o grupo das disciplinas, 1 excel das resoluções e 1 excel extraído do prime

Desenvolvido por Vinicius Tozo
Última atualização: 17/09/2021
'''

import re
from tkinter import Tk
from tkinter.filedialog import askopenfilename

import pandas
import xlrd
import os.path


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


def compara_matrizes(disciplinas_prime, disciplinas_resolucao, nome_arquivo, ordem_word):
    column_names = ["Coluna(s) com erro", "Ordem", "Ciclo", "Disciplina", "Requisitos", "Carga horária teórica",
                    "Carga horária prática", "Créditos", "Carga Horária oficial", "Carga Horária relógio oficial",
                    "Grupo"]

    root = "/".join(nome_arquivo.split("/")[0:-1]) + "/"
    curso = nome_arquivo.replace(root, '').replace('.xlsx', '').replace('Grade_curricular_', '').strip()
    nome_arquivo_saida = f"{root}Relatório de Erros {curso} (Resolução atualizada).xlsx"

    if os.path.isfile(nome_arquivo_saida):
        raise Exception("Arquivo já existe")

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

    columns = ["Ciclo", "Disciplina", "Requisitos", "Carga horária teórica", "Carga horária prática", "Créditos",
               "Carga Horária oficial", "Carga Horária relógio oficial", "Grupo"]

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
                grade_prime = disciplinas_prime[disciplinas_prime.index(indice)]

                for i, c in enumerate(columns):
                    if disciplina_resolucao[i] != grade_prime[i]:
                        erro.append(f"{c} \n\tResolução: {disciplina_resolucao[i]} \n\tPrime: {grade_prime[i]}")
            else:
                erro.append(f"Nome da disciplina (verificar todos os campos) \n\t"
                            f"Resolução: {disciplina_resolucao[1]} \n\tPrime: Disciplina não encontrada")
            separador_nome = " \n"
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


def ler_resolucao(caminho_resolucao, dict_grupos):
    workbook_word = xlrd.open_workbook(caminho_resolucao)
    try:
        sheet_word = workbook_word.sheet_by_name("Plan1")
    except:
        sheet_word = workbook_word.sheet_by_name("Planilha1")
    disciplinas_resolucao = []

    # Dados Word
    ciclo = 0
    ordem_word = []

    for linha in range(0, sheet_word.nrows):
        if "PERÍODO" in str.upper(str(sheet_word.cell_value(linha, 0))):
            ciclo = int(''.join(filter(str.isdigit, sheet_word.cell_value(linha, 0))))
        elif "PERÍODO" in str.upper(str(sheet_word.cell_value(linha, 1))):
            ciclo = int(''.join(filter(str.isdigit, sheet_word.cell_value(linha, 1))))
        try:
            int(sheet_word.cell_value(linha, 0).__str__().strip().strip(".0"))
        except ValueError:
            continue

        ordem_disciplina = int(sheet_word.cell_value(linha, 0).__str__().strip().strip(".0"))
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
            aula_pratica, creditos, ch_oficial, ch_relogio_oficial,
            dict_grupos.get(nome_disciplina, 'Disciplina não encontrada')])
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
        grupo = sheet_excel.cell_value(linha, 5).strip()

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
             ch_relogio_oficial, grupo])

    if not disciplinas_prime:
        raise Exception("Erro na leitura da matriz")

    return disciplinas_prime


def ler_grupos(caminho_grupos):
    # Cria um dicionário com os códigos de disciplinas
    try:
        df_grupo = pandas.read_excel(caminho_grupos, sheet_name='CRIAÇÃO DE MATRIZ PRESENCIAL', engine='openpyxl',
                                     skiprows=12)
    except ValueError:
        df_grupo = pandas.read_excel(caminho_grupos, sheet_name='Modelo Matriz Presencial', engine='openpyxl',
                                     skiprows=5)

    # Remove espaços dos nomes das colunas
    df_grupo.columns = [x.strip().upper() for x in df_grupo.columns]
    df_grupo.columns = ["DISCIPLINA" if x == "DISCIPLINAS" else x for x in df_grupo.columns]

    df_grupo = df_grupo[['DISCIPLINA', 'GRUPO DA DISCIPLINA']]
    dict_disciplina = {}
    for k, v in df_grupo.iterrows():
        dict_disciplina[str(v['DISCIPLINA']).strip()] = str(v['GRUPO DA DISCIPLINA']).strip()
    return dict_disciplina


def main():
    Tk().withdraw()

    caminho_grupos = askopenfilename(
        filetypes=[('Arquivo excel', '.xlsx')],
        title='Selecione o arquivo com as informações do grupo das disciplinas')
    dict_grupos = ler_grupos(caminho_grupos)

    caminho_resolucao = askopenfilename(
        filetypes=[('Arquivo excel', '.xlsx')],
        title='Selecione o arquivo da resolução')
    resolucao, ordem_word = ler_resolucao(caminho_resolucao, dict_grupos)

    caminho_prime = askopenfilename(
        filetypes=[('Arquivo excel', '.xlsx')],
        title='Selecione o arquivo extraído do Prime')
    grade_prime = ler_grade(caminho_prime, resolucao)

    compara_matrizes(grade_prime, resolucao, caminho_prime, ordem_word)


if __name__ == '__main__':
    main()
