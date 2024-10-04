import os
from tkinter import Tk, filedialog


def lista_matrizes(root):
    # Para cada diretório dentro do diretório raiz
    for dirpath, dirnames, filenames in os.walk(root):

        # Se o diretório possuir outros diretórios dentro dele, ignora
        if dirnames:
            continue

        if len(dirpath.replace('\\', '/').split('/')[-1]) <= 1:
            continue

        # Se não, conta como uma matriz e continua a contagem de arquivos
        print(f'------------------------------------------------------------------------------------')
        print(dirpath.replace('\\', '/').split('/')[-1])
        print('------------------------------------------------------------------------------------')

        ano_periodo, ano_periodo_faltante, ano_verifica = [], [], 2015

        for file in filenames:
            if 'TACH' in file and (file.endswith('.pdf') or file.endswith('.PDF')):
                file_year = file.split('20')[-1].split('.')
                if file_year[0] == '':
                    file_year[0] = '20'
                file_year = f'{file_year[0]}.{file_year[1]}'.split(' ')[0].split('-')[0].split('_')[0]
                ano_periodo.append(f'20{file_year}')

        # Verifica quais anos não estão presentes na
        while ano_verifica <= 2021:
            if str(ano_verifica) + '.1' not in ano_periodo:
                ano_periodo_faltante.append(f'{ano_verifica}.1')
            if str(ano_verifica) + '.2' not in ano_periodo:
                ano_periodo_faltante.append(f'{ano_verifica}.2')
            ano_verifica += 1

        # Remove os valores duplicados e ordena a lista
        if len(ano_periodo_faltante) > 0:
            ano_periodo_faltante = sorted(list(set(ano_periodo_faltante)))
        print()
        print('Faltam arquivos de', end=': ')
        for ap in ano_periodo_faltante:
            if ap == ano_periodo_faltante[-1]:
                print(f'{ap}\n')
                break
            print(ap, end=', ')
        # reseta as listas e o contador
        ano_periodo, ano_periodo_faltante, ano_verifica = [], [], 2015


if __name__ == '__main__':
    Tk().withdraw()
    pasta_raiz = filedialog.askdirectory()
    print()
    lista_matrizes(pasta_raiz)
