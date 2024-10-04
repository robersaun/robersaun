"""Script para Excluir Relatórios de Erros

Script para buscar nas pastas relatórios de erros para excluir.
Recebe como entrada a pasta raiz contendo todas as pastas com as matrizes / resoluções,
e exclui arquivos com o nome 'Relatório de Erros.xlsx' em pastas com a grade curricular e resolução, para que possa ser
gerado novamente.

Não é mais utilizado.


Desenvolvido por Vinicius Tozo
Última atualização: 13/08/2021
"""
import os
from tkinter import Tk, filedialog


def exclui_relatorios(root):
    # Para cada diretório dentro do diretório raiz
    for dirpath, dirnames, filenames in os.walk(root):

        # Se o diretório possuir outros diretórios dentro dele, ignora
        if dirnames:
            continue

        possui_resolucao = False
        possui_matriz = False

        # Para cada arquivo dentro do diretório, verifica se a matriz possui os arquivos necessários
        for file in filenames:
            if file.startswith("Grade_curricular_") and file.endswith(".xlsx"):
                possui_matriz = True
            elif file == "Novo(a) Planilha do Microsoft Excel.xlsx":
                possui_resolucao = True
            elif file == "Relatório de Erros.xlsx" and possui_matriz and possui_resolucao:
                arquivo_excluir = dirpath + "\\" + file
                print(arquivo_excluir)
                os.remove(arquivo_excluir)


if __name__ == '__main__':
    Tk().withdraw()
    pasta_raiz = filedialog.askdirectory()
    exclui_relatorios(pasta_raiz)
