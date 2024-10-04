"""Script para Encontrar Arquivos Faltantes

Script para buscar nas pastas quais arquivos ainda não foram criados.
Recebe como entrada a pasta raiz contendo todas as pastas com as matrizes / resoluções,
e imprime em tela quais pastas contém uma matriz e nenhum relatório de erro (esta verificação pode ser parametrizada),
e por fim imprime em tela uma contagem dos tipos de arquivo presentes nas pastas.

Pode ser utilizado pelo analista de TI para auxiliar no processo de sanitização de matrizes.


Desenvolvido por Vinicius Tozo
Última atualização: 24/09/2021
"""
import os
from tkinter import Tk, filedialog


def lista_matrizes(root):
    # Variáveis para a contagem
    contagem_total = 0
    contagem_matrizes = 0
    contagem_matrizes_v2 = 0
    contagem_resolucao_word = 0
    contagem_resolucao_excel = 0
    contagem_relatorio_erros = 0
    contagem_relatorio_erros_v2 = 0

    # Para cada diretório dentro do diretório raiz
    for dirpath, dirnames, filenames in os.walk(root):

        # Se o diretório possuir outros diretórios dentro dele, ignora
        if dirnames:
            continue

        # Se não, conta como uma matriz e continua a contagem de arquivos

        contagem_total += 1
        possui_matriz = False
        possui_matriz_v2 = False
        possui_resolucao_em_word = False
        possui_resolucao_em_excel = False
        possui_relatorio_erros = False
        possui_relatorio_erros_v2 = False

        # Para cada arquivo dentro do diretório, verifica se a matriz possui os arquivos necessários
        for file in filenames:
            if file.endswith(".xlsx"):
                possui_matriz = True
            if file.startswith("Relatório de Erros"):
                possui_relatorio_erros = True
            if "v2" in file or "V2" in file:
                possui_matriz_v2 = True
            if file.startswith("Relatório de Erros") and "(Resolução atualizada).xlsx" in file:
                possui_relatorio_erros_v2 = True

        if possui_matriz:
            contagem_matrizes += 1
        if possui_resolucao_em_word:
            contagem_resolucao_word += 1
        if possui_resolucao_em_excel:
            contagem_resolucao_excel += 1
        if possui_relatorio_erros:
            contagem_relatorio_erros += 1
        if possui_matriz_v2:
            contagem_matrizes_v2 += 1
        if possui_relatorio_erros_v2:
            contagem_relatorio_erros_v2 += 1

        if possui_matriz and not possui_relatorio_erros_v2:
            print(dirpath)

    # Imprime as contagens alinhadas à direita (>20)
    print()
    print(f"Matriz:                {str(contagem_matrizes) + '/' + str(contagem_total): >20}")
    print(f"Matriz v2:             {str(contagem_matrizes_v2) + '/' + str(contagem_total): >20}")
    print(f"Relatório de Erros:    {str(contagem_relatorio_erros) + '/' + str(contagem_total): >20}")
    print(f"Relatório de Erros v2: {str(contagem_relatorio_erros_v2) + '/' + str(contagem_total): >20}")


if __name__ == '__main__':
    Tk().withdraw()
    pasta_raiz = filedialog.askdirectory()
    print()
    lista_matrizes(pasta_raiz)
