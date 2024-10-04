'''
Cria um novo dicionario dos docentes e alunos usando o cpf como id e compara com os dados do ano anterior (excluindo os formandos)

Desenvolvido por Vinicius Tozo
Última atualização: 03/01/2021
'''

# -*- coding: utf-8 -*-
import json


def cria_dicionario_dados_atuais(arquivo):
    dicionario = {}
    txt = open(arquivo, encoding="windows-1252")

    identificador = ""
    for line in txt:
        line = line.strip()
        if line.startswith("31") or line.startswith("41"):  # Dados de docente ou aluno
            identificador = str(line.split("|")[2]) + " " + str(line.split("|")[3])  # Utilizando o nome e cpf como id
            if identificador not in dicionario:
                dicionario[identificador] = {"dados": line, "vinculos": []}
        elif line.startswith("32") or line.startswith("42"):  # Dados de vínculo
            dicionario[identificador]["vinculos"].append(line)
    return dicionario


def busca_dados_antigos(arquivo, dicionario):
    txt = open(arquivo, encoding="windows-1252")
    resultado = []
    identificador = ""
    for line in txt:
        line = line.strip()
        if line.startswith("31") or line.startswith("41"):  # Dados de docente ou aluno
            identificador = str(line.split("|")[2]) + " " + str(line.split("|")[3])  # Utilizando o nome e cpf como id
            if identificador not in dicionario:
                resultado.append(f"Dados não encontrados: {identificador}")
            elif dicionario[identificador]["dados"] != line:
                resultado.append(f"Dados pessoais não correspondem: {line}")
        elif line.startswith("32") or line.startswith("42"):  # Dados de vínculo
            if identificador not in dicionario:
                continue
            if line not in dicionario[identificador]["vinculos"]:
                resultado.append(f"Vínculo de {identificador} não corresponde: {line}")
    return resultado


if __name__ == '__main__':
    dicionario_dados = cria_dicionario_dados_atuais("Unificado_2020.txt")
    resultado_busca = busca_dados_antigos("Alunos_2019_sem_formados.txt", dicionario_dados)
    for r in resultado_busca:
        print(r)
    # print(json.dumps(dicionario_dados, sort_keys=True, indent=4))
