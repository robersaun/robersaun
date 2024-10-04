# -*- coding: Windows-1252 -*-

from tkinter import Tk
from tkinter.filedialog import askopenfilename
import pandas as pd


# noinspection PyUnboundLocalVariable,PyBroadException
# Retorna os valores da equival�ncia para fazer a busca e compara��o
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
            grade = grade[0].split('-')
            # Trata os casos em que h� ' -' ou ' -'
            # Evita que ocorra uma exce��o nos casos em que n�o h� turno
            try:
                matriz = grade[0].split(' ')
                turno = grade[1].split(' ')
            except:
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
            # Se houver a letra R ir� colocar o turno n�o tratado, visto que n�o precisa ser comparado
            try:
                if turno[-1] != 'R':
                    matriz = matriz+'-'+turno[-1]
                else:
                    matriz = matriz+'-'+grade[1]
            except:
                matriz = matriz
            ciclo = str(int(gcd[2].split('-')[0]))
            disciplina = disciplina[0].split(' ')[-1]
        # Ordem: Campo disciplina original, c�d disciplina, matriz, tipo de equival�ncia, ciclo
        return [eq[cont], disciplina, matriz, tipo[1]+tipo[2]+tipo[3], ciclo]
    except:
        return [eq[cont]]


# Retorna o resultado da compara��o das equival�ncias
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

    col_names = []
    for cn in df.columns:
        col_names.append(cn)

    data = {}
    valores = []

    for v in range(len(col_names)):
        valores.append([])

    # exit()

    # Separa as informa��es por linha
    for linha in range(df.shape[0]):
        if linha > 0:
            break
        # print(f'\n########################## Linha {linha+2} ##########################')
        eq = df["Equival�ncia(s)"][linha].splitlines()
        # Nova informa��o que ser� escrita nas equval�ncias

        # Para cada equival�ncia, faz as verifica��es
        for cont in range(len(eq)):
            # Ordem: Campo disciplina original, c�d disciplina, matriz, tipo de equival�ncia, ciclo
            dados = reescreve_equivalencia(eq, cont)
            equivalencia = dados[0]
            novo_texto = f'{equivalencia}'

            try:
                disciplina = dados[1]
                matriz = dados[2]

                for index, row in df.iterrows():
                    d_matriz = row[9]  # Coluna de matriz
                    d_disc = row[0]  # Coluna de disciplinas
                    if d_matriz == matriz and disciplina in d_disc:
                        # Ordem: Cr�dito, CH te�rica, CH pr�tica, CH oficial, CH rel�gio oficial
                        res_equivalencia = compara_equivalencia(df, linha, index)
                        # Atualiza o texto da equival�ncia
                        '''if novo_texto == '':
                            novo_texto = (f'{equivalencia} - Cr�dito: {res_equivalencia[0]} ' +
                                          f'- CH te�rica: {res_equivalencia[1]} ' +
                                          f'- CH pr�tica: {res_equivalencia[2]} ' +
                                          f'- CH oficial: {res_equivalencia[3]} ' +
                                          f'- CH rel�gio oficial: {res_equivalencia[4]}')
                        else:
                            novo_texto = novo_texto + (f'\n{equivalencia} - Cr�dito: {res_equivalencia[0]} ' +
                                                       f'- CH te�rica: {res_equivalencia[1]} ' +
                                                       f'- CH pr�tica: {res_equivalencia[2]} ' +
                                                       f'- CH oficial: {res_equivalencia[3]} ' +
                                                       f'- CH rel�gio oficial: {res_equivalencia[4]}')
                        break'''
                        novo_texto = (f'{equivalencia} - Cr�dito: {res_equivalencia[0]} ' +
                                      f'- CH te�rica: {res_equivalencia[1]} ' +
                                      f'- CH pr�tica: {res_equivalencia[2]} ' +
                                      f'- CH oficial: {res_equivalencia[3]} ' +
                                      f'- CH rel�gio oficial: {res_equivalencia[4]}')
                        break
            except:
                continue
            # print(dados)
            # print('---------------------------------------------------------------------------------------------')
            '''try:
                novo_texto = (f'{equivalencia} - Cr�dito: {res_equivalencia[0]} ' +
                              f'- CH te�rica: {res_equivalencia[1]} ' +
                              f'- CH pr�tica: {res_equivalencia[2]} ' +
                              f'- CH oficial: {res_equivalencia[3]} ' +
                              f'- CH rel�gio oficial: {res_equivalencia[4]}')
            except:
                continue'''
            valores[0].append(df[col_names[0]][linha])
            valores[1].append(df[col_names[1]][linha])
            valores[2].append(df[col_names[2]][linha])
            valores[3].append(df[col_names[3]][linha])
            valores[4].append(df[col_names[4]][linha])
            valores[5].append(df[col_names[5]][linha])
            valores[6].append(df[col_names[6]][linha])
            valores[7].append(df[col_names[7]][linha])
            valores[8].append(df[col_names[8]][linha])
            valores[9].append(df[col_names[9]][linha])
            valores[10].append(novo_texto)
        # valores[10].append(novo_texto)
        # print(novo_texto)
        # print(f'########################## Linha {linha+2} ##########################')

    '''print(valores)
    print(len(valores[0]))
    print(len(valores[1]))
    print(len(valores[2]))
    print(len(valores[3]))
    print(len(valores[4]))
    print(len(valores[5]))
    print(len(valores[6]))
    print(len(valores[7]))
    print(len(valores[8]))
    print(len(valores[9]))
    print(len(valores[10]))
    for pp in range(len(valores[10])):
        print(f'{pp}: {valores[10][pp]}')
    exit()'''
    # TODO: Corrigir o erro
    # TypeError: list indices must be integers or slices, not str
    '''for col in range(len(col_names)):
        try:
            nome_col = col_names[col]
            dados_coluna = valores[col]
            dados[nome_col] = []
            dados[nome_col].append(dados_coluna)
            # dados[col] = valores[col]
        except:
            break'''
    n_df = pd.DataFrame(data={col_names[0]: valores[0],
                              col_names[1]: valores[1],
                              col_names[2]: valores[2],
                              col_names[3]: valores[3],
                              col_names[4]: valores[4],
                              col_names[5]: valores[5],
                              col_names[6]: valores[6],
                              col_names[7]: valores[7],
                              col_names[8]: valores[8],
                              col_names[9]: valores[9],
                              col_names[10]: valores[10]})
    # print(n_df)
    n_df.to_excel('equivalencias.xlsx', index=False)
    print('Excel gerado')


if __name__ == '__main__':
    # Chamada da fun��o main
    main()
