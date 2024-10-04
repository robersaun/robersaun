"""Gerador de Abreviação para Disciplinas

Script para gerar o abreviações únicas de nomes de disciplinas para usar no sistema PowerCubus.
Recebe como entrada o arquivo txt 'disciplinas.txt' que deve estar na mesma pasta com a lista de nomes.
Imprime em tela a lista de abreviações na mesma ordem do txt de entrada.

Pode ser utilizado pelo analista de TI para auxiliar no processo de elaboração de horários no PowerCubus.

Desenvolvido por Vinicius Tozo
Última atualização: 16/07/2021
"""
import random
import string


def main():
    dicionario = {}
    disciplinas = open("disciplinas.txt", "r", encoding="utf-8").read()

    for disciplina in disciplinas.split("\n"):
        # Remove palavras que começam com lowercase
        palavras = " ".join(filter(lambda x: not x[0].islower(), disciplina.split(" ")))
        palavras = palavras.split(" ")

        num_palavras = len(palavras)
        letras_por_palavra = int(10 / num_palavras)

        sigla = ""
        for palavra in palavras:
            sigla += palavra[0:letras_por_palavra]

        while sigla in dicionario:
            sigla = sigla[:-1] + random.choice(string.ascii_lowercase)

        dicionario[sigla] = disciplina

    for key, value in dicionario.items():
        print(key)


if __name__ == '__main__':
    main()
