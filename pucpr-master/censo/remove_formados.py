'''
Cria um dicionário com os alunos e realiza a retira dos alunos formandos dele

Desenvolvido por Vinicius Tozo
Última atualização: 03/01/2021

'''

# -*- coding: utf-8 -*-
import json


def cria_dicionario(arquivo):
    dicionario = {}
    txt = open(arquivo, encoding="windows-1252")

    identificador = ""
    for line in txt:
        if line.startswith("31") or line.startswith("41"):  # Dados de docente ou aluno
            identificador = str(line.split("|")[2])  # Utilizando o nome como id
            dicionario[identificador] = {"dados": line, "vinculos": []}
        elif line.startswith("32") or line.startswith("42"):  # Dados de vínculo
            if line.split("|")[6] != "4" and line.split("|")[6] != "6":
                dicionario[identificador]["vinculos"].append(line)
    return dicionario


if __name__ == '__main__':
    dicionario_dados = cria_dicionario("Alunos_2019.txt")
    dicionario_dados = {k: v for k, v in dicionario_dados.items() if v["vinculos"] != []}  # Remove alunos sem vínculo
    for entry in dicionario_dados:
        print(dicionario_dados[entry]["dados"].strip())
        for vinculo in dicionario_dados[entry]["vinculos"]:
            print(vinculo.strip())
