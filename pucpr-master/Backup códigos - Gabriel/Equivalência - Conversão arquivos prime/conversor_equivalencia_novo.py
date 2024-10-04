# -*- coding: Windows-1252 -*-

from tkinter import Tk, filedialog
import pandas as pd
import os


# Verifica o nome das coluna e ajusta para os que precisa
def recupera_colunas(df):
    col_names = ['Ciclo']
    for cn in df.columns:
        col_names.append(cn)

    # Ajusta para ficar conforme o que ser� usado
    # col_names.pop()
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
    ch_t = int(df['C.H. Te�rica'][linha])
    ch_t_r = int(df['C.H. Te�rica'][index])
    if ch_t > ch_t_r:
        ch_t_eq = 'menor'
    elif ch_t < ch_t_r:
        ch_t_eq = 'maior'
    else:
        ch_t_eq = 'igual'

    # Carga hor�ria pr�tica
    ch_p = int(df['C.H. Pr�tica'][linha])
    ch_p_r = int(df['C.H. Pr�tica'][index])
    if ch_p > ch_p_r:
        ch_p_eq = 'menor'
    elif ch_p < ch_p_r:
        ch_p_eq = 'maior'
    else:
        ch_p_eq = 'igual'

    # Carga hor�ria oficial
    ch_o = int(df['C.H. Oficial'][linha])
    ch_o_r = int(df['C.H. Oficial'][index])
    if ch_o > ch_o_r:
        ch_o_eq = 'menor'
    elif ch_o < ch_o_r:
        ch_o_eq = 'maior'
    else:
        ch_o_eq = 'igual'

    # Carga hor�ria rel�gio oficial
    ch_ro = int(df['C.H. Rel�gio Oficial'][linha])
    ch_ro_r = int(df['C.H. Rel�gio Oficial'][index])
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
                                col_names[18]: valores[18],
                                col_names[19]: valores[19],
                                col_names[20]: valores[20],
                                col_names[21]: valores[21],
                                col_names[22]: valores[22],
                                col_names[23]: valores[23],
                                col_names[24]: valores[24],
                                col_names[25]: valores[25],
                                col_names[26]: valores[26],
                                col_names[27]: valores[27],
                                col_names[28]: valores[28],
                                col_names[29]: valores[29]
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
def armazena_valores(valores_b, df, col_names, linha, dados, res_equivalencia, ciclo):
    valores = valores_b
    try:
        for p in range(len(valores)):
            if p == 0:
                valores[p].append(ciclo)
            elif p < 21:
                valores[p].append(df[col_names[p]][linha])
            elif p < 25:
                valores[p].append(dados[p-20])
            else:
                valores[p].append(res_equivalencia[p-25])
        return valores
    except:
        return valores


# Fun��o principal
# noinspection PyBroadException
def main(root):
    # Para cada diret�rio dentro do diret�rio raiz
    for dirpath, dirnames, planilha in os.walk(root):

        # Se o diret�rio possuir outros diret�rios dentro dele, ignora
        if dirnames:
            continue
        # Sele��o do arquivo

        curso = dirpath.replace('\\', '/').split('/')[-1]
        os.makedirs(curso, exist_ok=True)

        # nome_arquivo = str(planilha).split('/')[-1]
        # Mostra qual foi o arquivo selecionado
        for nome_arquivo in planilha:
            print(f"Arquivo selecionado: {nome_arquivo}")
            doc_grade = f"{nome_arquivo.split('_')[-2]} {nome_arquivo.split('_')[-1].split('.')[0]}"

            ''' -------------------------------------------------------------------------------
                Leitura das p�ginas do arquivo
                ------------------------------------------------------------------------------- '''

            # Planilha de equival�ncias � armazenada na vari�vel
            # Os dados dos alunos dever�o sempre estar na primeira p�gina do arquivo
            arquivo = f'{dirpath}/{nome_arquivo}'
            df = pd.read_excel(arquivo, dtype=str, sheet_name='Ciclos', skiprows=5, usecols=lambda x: 'Unnamed' not in x)
            df.dropna(inplace=True, how='all')
            df = df.fillna("-")
            # Pega o nome das colunas
            col_names = recupera_colunas(df)

            valores = []
            for v in range(len(col_names)):
                valores.append([])

            ciclo = 1
            # Separa as informa��es por linha
            for linha in range(df.shape[0]):
                try:
                    c = int(df['Classif.'][linha])
                    if c == ciclo + 1:
                        ciclo = c
                except:
                    pass
                try:
                    eq = df["Equival�ncia(s)"][linha].split('<br />')
                except:
                    continue
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
                    valores = armazena_valores(valores, df, col_names, linha, dados, res_equivalencia, ciclo)
            print('Gerando Excel...')
            n_df = novo_df(col_names, valores)
            # Remove a coluna para tirar a redund�ncia
            n_df.pop('Equival�ncia(s)')
            n_df.to_excel(f'{curso}/equivalencias_{doc_grade}.xlsx', index=False)
            print('Excel gerado')


if __name__ == '__main__':
    # Chamada da fun��o main
    Tk().withdraw()
    pasta_raiz = filedialog.askdirectory()
    main(pasta_raiz)

#TODO: Add o Ciclo
#TODO: Arrumar para pegar os valores direito - OK? Oqq eu devia fazer aqui?
#TODO: Alterar a gera��o do nome do arquivo - OK
#TODO: Como que vou comparar as cargas hor�rias das equival�ncias?

# Aparentemnte agr � s� selecionar o diret�rio raiz q roda td