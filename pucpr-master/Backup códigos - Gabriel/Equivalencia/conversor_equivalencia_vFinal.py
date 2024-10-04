# -*- coding: Windows-1252 -*-

from tkinter import Tk
from tkinter.filedialog import askopenfilename
import pandas as pd


# Verifica o nome das coluna e ajusta para os que precisa
def recupera_colunas(df):
    col_names = []
    for cn in df.columns:
        col_names.append(cn)

    # Ajusta para ficar conforme o que será usado
    col_names.pop()
    novas_colunas = ['Tipo de equivalência', 'Grade da equivalência', 'Ciclo da equivalência',
                     'Disciplina da equivalência', 'Créditos da equivalência', 'CH Teórica da equivalência',
                     'CH Prática da equivalência', 'CH Oficial da equivalência',
                     'CH Relógio Oficial da equivalência']

    for nc in novas_colunas:
        col_names.append(nc)
    return col_names


# Retorna os valores da equivalência para fazer a busca e comparação
# noinspection PyUnboundLocalVariable,PyBroadException
def dados_equivalencia(eq, cont):
    try:
        # Tira a barrinha
        equivalencia = eq[cont].split('|')
        # Tira os dois pontinhos do tipo de quivalência
        tipo = equivalencia[0].split(':')[-1].split(' ')
        # Tira os dois pontinhos da grade, ciclo e disciplina
        gcd = equivalencia[-1].split(':')
        # Seleciona as partes que dizem a matriz e a disciplina
        grade = gcd[1].split(' - ')
        disciplina = gcd[3].split(' - ')
        # Pegar o valor para comparar
        grade = grade[0].split('-')
        # Trata os casos em que há ' -' ou ' -'
        # Evita que ocorra uma exceção nos casos em que não há turno
        try:
            matriz = grade[0].split(' ')
            turno = grade[1].split(' ')
        except:
            matriz = grade[0].split(' ')
        try:
            # Trata espaços em branco e os casos m que não há '-TURNO'
            if matriz[3] != '':
                matriz = matriz[1] + ' ' + matriz[2] + ' ' + matriz[3]
            else:
                matriz = matriz[1] + ' ' + matriz[2]
        except:
            matriz = matriz[1] + ' ' + matriz[2]
        # Tenta inserir o turno
        # Se houver a letra R irá colocar o turno não tratado, visto que não precisa ser comparado
        try:
            if turno[-1] != 'R':
                matriz = matriz + '-' + turno[-1]
            else:
                matriz = matriz + '-' + grade[1]
        except:
            matriz = matriz
        ciclo = str(int(gcd[2].split('-')[0]))
        disciplina = disciplina[0].split(' ')[-1]
        # Ordem: Campo disciplina original, tipo de equivalência, matriz, ciclo, cód disciplina
        return [eq[cont], tipo[1] + ' ' + tipo[2] + ' ' + tipo[3], matriz, ciclo, disciplina]
    except:
        return [eq[cont], '', '', '', '']


# Retorna o resultado da comparação das equivalências
def compara_equivalencia(df, linha, index, tipo):
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


# Verifica, quando possível, se a equivalencia é menor, igual ou maior
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
    # Para tratar espaços em branco
    disciplina = disciplina.split(': ')
    try:
        disciplina = f'{disciplina[3]}: {disciplina[4]}'
    except:
        disciplina = disciplina[3]
    return disciplina


# Armazena os valores que serão inseridos na tabela
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


# Função principal
# noinspection PyBroadException
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

    # Pega o nome das colunas
    col_names = recupera_colunas(df)

    valores = []
    for v in range(len(col_names)):
        valores.append([])

    # Separa as informações por linha
    for linha in range(df.shape[0]):
        eq = df["Equivalência(s)"][linha].splitlines()

        # Para cada equivalência, faz as verificações
        for cont in range(len(eq)):
            # Ordem: Campo disciplina original, tipo de equivalência, matriz, ciclo, cód disciplina
            dados = dados_equivalencia(eq, cont)
            try:
                dados[4] = nome_disciplina(eq, cont)
            except:
                continue
            # Ordem: Crédito, CH teórica, CH prática, CH oficial, CH relógio oficial
            res_equivalencia = resultado_equivalencia(df, linha, dados)
            valores = armazena_valores(valores, df, col_names, linha, dados, res_equivalencia)
    print('Gerando Excel...')
    n_df = novo_df(col_names, valores)
    n_df.to_excel('equivalencias_verificação.xlsx', index=False)
    print('Excel gerado')


if __name__ == '__main__':
    # Chamada da função main
    main()
