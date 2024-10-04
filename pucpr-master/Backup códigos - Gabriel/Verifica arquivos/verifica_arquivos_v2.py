# -*- coding: Windows-1252 -*-
from tkinter.filedialog import askopenfilename
import pandas as pd
import os
from tkinter import Tk, filedialog


def lista_matrizes(root):

    lista_profs = []
    lista_pendencias = []

    # Seleção do arquivo
    Tk().withdraw()
    planilha = askopenfilename(
        filetypes=[('Arquivo excel', '.xlsx')],
        title='Selecione a planilha de admissões')

    df = pd.read_excel(planilha, dtype=str, sheet_name=0, header=3, usecols='C:D')
    df = df.fillna("")

    # Para cada diretório dentro do diretório raiz
    for dirpath, dirnames, filenames in os.walk(root):

        # Se o diretório possuir outros diretórios dentro dele, ignora
        if dirnames:
            continue

        prof = dirpath.replace('\\', '/').split('/')[-1]
        if len(prof) <= 1:
            continue
        # Busca o prof e a data de admissão
        index = df.index[df['NOME'] == prof].tolist()[0]
        ano_mes_admissao = str(df['ADMISSÃO'][index].split()[0])

        ano_admissao = int(ano_mes_admissao.split('-')[0])
        mes_admissao = int(ano_mes_admissao.split('-')[1])

        lista_profs.append(prof)
        ano_periodo, ano_periodo_faltante = [], []

        # Verifica se o ano de admissão do prof, para não cobrar os documentos desnecessários
        if ano_admissao > 2015:
            ano_verifica = ano_admissao
            if 1 <= mes_admissao < 6:
                semestre = 1
            else:
                semestre = 2
        else:
            ano_verifica = 2015
            # Como quem entra no else já estava admitido previamente, o semestre fica como 1 independente do mês
            semestre = 1

        for file in filenames:
            if 'TACH' in file and (file.endswith('.pdf') or file.endswith('.PDF')):
                ano_periodo.append(file.replace("-", "."))

        # Verifica quais anos não estão presentes na
        while ano_verifica <= 2021:
            if (str(ano_verifica) + '.2' == str(ano_admissao) + '.' + str(semestre) and
                    str(ano_verifica) + '.2' not in str(ano_periodo)):
                ano_periodo_faltante.append(f'{ano_verifica}.2')
            else:
                if str(ano_verifica) + '.1' not in str(ano_periodo):
                    ano_periodo_faltante.append(f'{ano_verifica}.1')
                if str(ano_verifica) + '.2' not in str(ano_periodo):
                    ano_periodo_faltante.append(f'{ano_verifica}.2')
            ano_verifica += 1

        # Remove os valores duplicados e ordena a lista
        if len(ano_periodo_faltante) > 0:
            ano_periodo_faltante = sorted(list(set(ano_periodo_faltante)))
        ano_periodo_pendente = ''
        for ap in ano_periodo_faltante:
            if ap == ano_periodo_faltante[-1]:
                ano_periodo_pendente += f'{ap}'
                break
            ano_periodo_pendente += f'{ap}, '
        lista_pendencias.append(ano_periodo_pendente)

    dados = {'Professor/a': lista_profs, 'Documentos Pendentes': lista_pendencias}
    ndf = pd.DataFrame(data=dados)
    ndf.to_excel('Arquivos_pendentes_profs.xlsx', index=False, sheet_name='Pendências')


if __name__ == '__main__':
    Tk().withdraw()
    pasta_raiz = filedialog.askdirectory()
    print()
    lista_matrizes(pasta_raiz)
