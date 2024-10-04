# -*- coding: Windows-1252 -*-

from tkinter import Tk
from tkinter.filedialog import askopenfilename
import pandas as pd


# noinspection PyUnboundLocalVariable,PyBroadException
def reescreve_equivalencia(eq, cont):
    try:
        # Tira a barrinha
        equivalencia = eq[cont].split('|')
        # Tira os dois pontinhos do tipo de quival�ncia
        tipo = equivalencia[0].split(':')[-1].split(' ')
        v1 = tipo[1]
        v2 = tipo[3]
        # Verifica se a equival�ncia � 1 para 1
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
            grade = grade[0].split('-')  # Trata os casos em que h� '- ' ou ' -'
            matriz = grade[0].split(' ')
            try:
                # Trata espa�os em branco e os casos m que n�o h� '-TURNO'
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
    # Cr�dito
    cr = int(df['Cr�ditos'][linha])
    cr_r = int(df['Cr�ditos'][index])
    if cr > cr_r:
        cr_eq = 'menor'
    elif cr < cr_r:
        cr_eq = 'maior'
    else:
        cr_eq = 'igual'

    # Carga hor�ria te�rica
    ch_t = int(df['CH Te�rica'][linha])
    ch_t_r = int(df['CH Te�rica'][index])
    if ch_t > ch_t_r:
        ch_t_eq = 'menor'
    elif ch_t < ch_t_r:
        ch_t_eq = 'maior'
    else:
        ch_t_eq = 'igual'

    # Carga hor�ria pr�tica
    ch_p = int(df['CH Pr�tica'][linha])
    ch_p_r = int(df['CH Pr�tica'][index])
    if ch_p > ch_p_r:
        ch_p_eq = 'menor'
    elif ch_p < ch_p_r:
        ch_p_eq = 'maior'
    else:
        ch_p_eq = 'igual'

    # Carga hor�ria oficial
    ch_o = int(df['CH Oficial'][linha])
    ch_o_r = int(df['CH Oficial'][index])
    if ch_o > ch_o_r:
        ch_o_eq = 'menor'
    elif ch_o < ch_o_r:
        ch_o_eq = 'maior'
    else:
        ch_o_eq = 'igual'

    # Carga hor�ria rel�gio oficial
    ch_ro = int(df['CH Rel�gio Oficial'][linha])
    ch_ro_r = int(df['CH Rel�gio Oficial'][index])
    if ch_ro > ch_ro_r:
        ch_ro_eq = 'menor'
    elif ch_ro < ch_ro_r:
        ch_ro_eq = 'maior'
    else:
        ch_ro_eq = 'igual'

    # Ordem: Cr�dito, CH te�rica, CH pr�tica, CH oficial, CH rel�gio oficial
    return [f'{cr_r} ({cr_eq})', f'{ch_t_r} ({ch_t_eq})', f'{ch_p_r} ({ch_p_eq})',
            f'{ch_o_r} ({ch_o_eq})', f'{ch_ro_r} ({ch_ro_eq})']


# Fun��o principal
def main():
    # Sele��o do arquivo
    print("Selecione a planilha de alunos")
    Tk().withdraw()
    planilha = askopenfilename(
        filetypes=[('Arquivo excel', '.xlsx')],
        title='Selecione a planilha de alunos')

    # Mostra qual foi o arquivo selecionado
    print("Arquivo selecionado: " + str(planilha).split('/')[-1])

    ''' -------------------------------------------------------------------------------
        Leitura das p�ginas do arquivo
        ------------------------------------------------------------------------------- '''

    print("Lendo arquivo")

    # Planilha de equival�ncias � armazenada na vari�vel
    # Os dados dos alunos dever�o sempre estar na primeira p�gina do arquivo
    df = pd.read_excel(planilha, dtype=str, sheet_name=0)

    df = df.fillna("")

    # Separa as informa��es por linha
    for linha in range(df.shape[0]):
        if linha > 0:
            break
        # print(f'\n########################## Linha {linha+2} ##########################')
        eq = df["Equival�ncia(s)"][linha].splitlines()
        # Nova informa��o que ser� escrita nas equval�ncias
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
    # Chamada da fun��o main
    main()
