"""Gerador de Abreviação para Nomes de Professores

Script para gerar o abreviações únicas de nomes de professores para usar no sistema PowerCubus.
Recebe como entrada o arquivo txt 'professores.txt' que deve estar na mesma pasta com a lista de nomes.
Imprime em tela a lista de abreviações na mesma ordem do txt de entrada.

Pode ser utilizado pelo analista de TI para auxiliar no processo de elaboração de horários no PowerCubus.

Desenvolvido por Vinicius Tozo
Última atualização: 19/10/2021
"""


def main():
    dicionario = {}
    professores = open("professores.txt", "r", encoding="utf-8").read()

    for nome in professores.split("\n"):

        palavras = ' '.join(filter(lambda x: not x[0].islower(), nome.strip().split(' ')))
        palavras = palavras.split(' ')

        # Pega as iniciais de cada palavra e o máximo possível da primeira palavra sem
        # ultrapassar o limite de 10 caracteres
        # Exemplo: 'Leitura e Escrita de Textos Técnico-Científicos' fica 'Leitu ETTC'
        sigla = ''
        for palavra in palavras[1:]:
            sigla += palavra[0]
        sigla = palavras[0][0:10 - len(palavras)] + ' ' + sigla

        if sigla in dicionario.values():
            print(f'Colisão detectada: {nome} => {sigla}')

        dicionario[nome] = sigla

    for key, value in dicionario.items():
        print(value)


if __name__ == '__main__':
    main()
