# -*- coding: Windows-1252 -*-

from tkinter import Tk
from tkinter.filedialog import askopenfilename
import pandas as pd


# Verifica o nome das coluna e ajusta para os que precisa
def recupera_colunas(df):
    col_names = []
    for cn in df.columns:
        col_names.append(cn)

    # Ajusta para ficar conforme o que ser� usado
    col_names.pop()
    novas_colunas = ['Tipo de equival�ncia', 'Grade da equival�ncia', 'Ciclo da equival�ncia',
                     'Disciplina da equival�ncia', 'Cr�ditos da equival�ncia', 'CH Te�rica da equival�ncia',
                     'CH Pr�tica da equival�ncia', 'CH Oficial da equival�ncia',
                     'CH Rel�gio Oficial da equival�ncia']

    for nc in novas_colunas:
        col_names.append(nc)
    return col_names


# Retorna os valores da equival�ncia para fazer a busca e compara��o
# noinspection PyUnboundLocalVariable,PyBroadException
def dados_equivalencia(eq, cont):
    try:
        # Tira a barrinha
        equivalencia = eq[cont].split('|')
        # Tira os dois pontinhos do tipo de quival�ncia
        tipo = equivalencia[0].split(':')[-1].split(' ')
        # Tira os dois pontinhos da grade, ciclo e disciplina
        gcd = equivalencia[-1].split(':')
        # Seleciona as partes que dizem a matriz e a disciplina
        grade = gcd[1].split(' - ')
        disciplina = gcd[3].split(' - ')
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
                matriz = matriz[1] + ' ' + matriz[2] + ' ' + matriz[3]
            else:
                matriz = matriz[1] + ' ' + matriz[2]
        except:
            matriz = matriz[1] + ' ' + matriz[2]
        # Tenta inserir o turno
        # Se houver a letra R ir� colocar o turno n�o tratado, visto que n�o precisa ser comparado
        try:
            if turno[-1] != 'R':
                matriz = matriz + '-' + turno[-1]
            else:
                matriz = matriz + '-' + grade[1]
        except:
            matriz = matriz
        ciclo = str(int(gcd[2].split('-')[0]))
        disciplina = disciplina[0].split(' ')[-1]
        # Ordem: Campo disciplina original, tipo de equival�ncia, matriz, ciclo, c�d disciplina
        return [eq[cont], tipo[1] + ' ' + tipo[2] + ' ' + tipo[3], matriz, ciclo, disciplina]
    except:
        return [eq[cont], '', '', '', '']


# Retorna o resultado da compara��o das equival�ncias
def compara_equivalencia(df, linha, index, tipo):
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
    if 'N' in tipo or 'n' in tipo:
        return [f'{cr_r}', f'{ch_t_r}', f'{ch_p_r}', f'{ch_o_r}', f'{ch_ro_r}']
    else:
        return [f'{cr_r} ({cr_eq})', f'{ch_t_r} ({ch_t_eq})', f'{ch_p_r} ({ch_p_eq})',
                f'{ch_o_r} ({ch_o_eq})', f'{ch_ro_r} ({ch_ro_eq})']


# Insere os valores no dataframe
def novo_df(col_names, valores):
    new_df = pd.DataFrame(data={col_names[0]: valores[0],
                                col_names[1]: valores[1],
                                col_names[2]: valores[2],
                                col_names[3]: valores[3],
                                col_names[4]: valores[4],
                                col_names[5]: valores[5],
                                col_names[6]: valores[6],
                                col_names[7]: valores[7],
                                col_names[8]: valores[8],
                                col_names[9]: valores[9],
                                col_names[10]: valores[10],
                                col_names[11]: valores[11],
                                col_names[12]: valores[12],
                                col_names[13]: valores[13],
                                col_names[14]: valores[14],
                                col_names[15]: valores[15],
                                col_names[16]: valores[16],
                                col_names[17]: valores[17],
                                col_names[18]: valores[18]
                                })
    return new_df


# Verifica, quando poss�vel, se a equivalencia � menor, igual ou maior
# noinspection PyBroadException
def resultado_equivalencia(df, linha, dados):
    try:
        matriz = dados[2]
        disciplina = dados[4]
        tipo = dados[1]

        index = df.index[(df['Matriz'] == matriz) & (df['Disciplina'] == disciplina)].tolist()[0]
        res_equivalencia = compara_equivalencia(df, linha, index, tipo)
        return res_equivalencia
    except:
        return ['', '', '', '', '']


# Ajusta o nome da disciplina para inserir na tabela
# noinspection PyBroadException
def nome_disciplina(eq, cont):
    # Tira a barrinha
    equivalencia = eq[cont].split('|')
    disciplina = equivalencia[-1]
    # Para tratar espa�os em branco
    disciplina = disciplina.split(': ')
    try:
        disciplina = f'{disciplina[3]}: {disciplina[4]}'
    except:
        disciplina = disciplina[3]
    return disciplina


# Armazena os valores que ser�o inseridos na tabela
def armazena_valores(valores_b, df, col_names, linha, dados, res_equivalencia):
    valores = valores_b
    for p in range(len(valores)):
        if p < 10:
            valores[p].append(df[col_names[p]][linha])
        elif p < 14:
            valores[p].append(dados[p-9])
        else:
            valores[p].append(res_equivalencia[p-14])
    return valores


# Fun��o principal
# noinspection PyBroadException
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

    # Pega o nome das colunas
    col_names = recupera_colunas(df)

    valores = []
    for v in range(len(col_names)):
        valores.append([])

    # Separa as informa��es por linha
    for linha in range(df.shape[0]):
        eq = df["Equival�ncia(s)"][linha].splitlines()

        # Para cada equival�ncia, faz as verifica��es
        for cont in range(len(eq)):
            # Ordem: Campo disciplina original, tipo de equival�ncia, matriz, ciclo, c�d disciplina
            dados = dados_equivalencia(eq, cont)
            try:
                dados[4] = nome_disciplina(eq, cont)
            except:
                continue
            # Ordem: Cr�dito, CH te�rica, CH pr�tica, CH oficial, CH rel�gio oficial
            res_equivalencia = resultado_equivalencia(df, linha, dados)
            valores = armazena_valores(valores, df, col_names, linha, dados, res_equivalencia)
    print('Gerando Excel...')
    n_df = novo_df(col_names, valores)
    n_df.to_excel('equivalencias_verifica��o.xlsx', index=False)
    print('Excel gerado')


if __name__ == '__main__':
    # Chamada da fun��o main
    main()
