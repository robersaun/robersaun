# -*- coding: Windows-1252 -*-

from tkinter import Tk
from tkinter.filedialog import askopenfilename
import pandas as pd


# noinspection PyUnboundLocalVariable,PyBroadException
def reescreve_equivalencia(eq, cont):
    try:
        # Tira a barrinha
        equivalencia = eq[cont].split('|')
        # Tira os dois pontinhos do tipo de quivalência
        tipo = equivalencia[0].split(':')[-1].split(' ')
        v1 = tipo[1]
        v2 = tipo[3]
        # Verifica se a equivalência é 1 para 1
        if v1 == v2:
            igual = True
        else:
            igual = False
        if igual:
            # Tira os dois pontinhos da grade, ciclo e disciplina
            gcd = equivalencia[-1].split(':')
            # Seleciona as partes que dizem a matriz e a disciplina
            grade = gcd[1].split(' - ')
            disciplina = gcd[-1].split(' - ')
            # Pegar o valor para comparar
            grade = grade[0].split('-')  # Trata os casos em que há '- ' ou ' -'
            matriz = grade[0].split(' ')
            try:
                # Trata espaços em branco e os casos m que não há '-TURNO'
                if matriz[3] != '':
                    matriz = matriz[1]+' '+matriz[2]+' '+matriz[3]
                else:
                    matriz = matriz[1] + ' ' + matriz[2]
            except:
                matriz = matriz[1]+' '+matriz[2]
            # Tenta inserir o turno
            try:
                matriz = matriz+'-'+grade[1]
            except:
                matriz = matriz
            disciplina = disciplina[0].split(' ')[-1]
        return [eq[cont], disciplina, matriz]
    except:
        return [eq[cont]]


def compara_equivalencia(df, linha, index):
    # Crédito
    cr = int(df['Créditos'][linha])
    cr_r = int(df['Créditos'][index])
    if cr > cr_r:
        cr_eq = 'menor'
    elif cr < cr_r:
        cr_eq = 'maior'
    else:
        cr_eq = 'igual'

    # Carga horária teórica
    ch_t = int(df['CH Teórica'][linha])
    ch_t_r = int(df['CH Teórica'][index])
    if ch_t > ch_t_r:
        ch_t_eq = 'menor'
    elif ch_t < ch_t_r:
        ch_t_eq = 'maior'
    else:
        ch_t_eq = 'igual'

    # Carga horária prática
    ch_p = int(df['CH Prática'][linha])
    ch_p_r = int(df['CH Prática'][index])
    if ch_p > ch_p_r:
        ch_p_eq = 'menor'
    elif ch_p < ch_p_r:
        ch_p_eq = 'maior'
    else:
        ch_p_eq = 'igual'

    # Carga horária oficial
    ch_o = int(df['CH Oficial'][linha])
    ch_o_r = int(df['CH Oficial'][index])
    if ch_o > ch_o_r:
        ch_o_eq = 'menor'
    elif ch_o < ch_o_r:
        ch_o_eq = 'maior'
    else:
        ch_o_eq = 'igual'

    # Carga horária relógio oficial
    ch_ro = int(df['CH Relógio Oficial'][linha])
    ch_ro_r = int(df['CH Relógio Oficial'][index])
    if ch_ro > ch_ro_r:
        ch_ro_eq = 'menor'
    elif ch_ro < ch_ro_r:
        ch_ro_eq = 'maior'
    else:
        ch_ro_eq = 'igual'

    # Ordem: Crédito, CH teórica, CH prática, CH oficial, CH relógio oficial
    return [f'{cr_r} ({cr_eq})', f'{ch_t_r} ({ch_t_eq})', f'{ch_p_r} ({ch_p_eq})',
            f'{ch_o_r} ({ch_o_eq})', f'{ch_ro_r} ({ch_ro_eq})']


# Função principal
def main():
    # Seleção do arquivo
    print("Selecione a planilha de alunos")
    Tk().withdraw()
    planilha = askopenfilename(
        filetypes=[('Arquivo excel', '.xlsx')],
        title='Selecione a planilha de alunos')

    # Mostra qual foi o arquivo selecionado
    print("Arquivo selecionado: " + str(planilha).split('/')[-1])

    ''' -------------------------------------------------------------------------------
        Leitura das páginas do arquivo
        ------------------------------------------------------------------------------- '''

    print("Lendo arquivo")

    # Planilha de equivalências é armazenada na variável
    # Os dados dos alunos deverão sempre estar na primeira página do arquivo
    df = pd.read_excel(planilha, dtype=str, sheet_name=0)

    df = df.fillna("")

    # Separa as informações por linha
    for linha in range(df.shape[0]):
        if linha > 0:
            break
        # print(f'\n########################## Linha {linha+2} ##########################')
        eq = df["Equivalência(s)"][linha].splitlines()
        # Nova informação que será escrita nas equvalências
        novo_texto = ''

        for cont in range(len(eq)):
            dados = reescreve_equivalencia(eq, cont)
            equivalencia = dados[0]
            try:
                disciplina = dados[1]
                matriz = dados[2]

                for index, row in df.iterrows():
                    d_matriz = row[9]
                    d_disc = row[0]
                    if d_matriz == matriz and disciplina in d_disc:
                        res_equivalencia = compara_equivalencia(df, linha, index)
                        print(res_equivalencia)
                        break
            except:
                continue
            # print(dados)
            # print('---------------------------------------------------------------------------------------------')
        print(novo_texto)
        # print(f'########################## Linha {linha+2} ##########################')


if __name__ == '__main__':
    # Chamada da função main
    main()
