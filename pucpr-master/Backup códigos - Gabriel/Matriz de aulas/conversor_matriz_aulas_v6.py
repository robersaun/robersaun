# -*- coding: Windows-1252 -*-

from tkinter import Tk
from tkinter.filedialog import askopenfilename
from numpy import nan
import pandas as pd

def ajusta_info(df_aulas):
    cont = 1 # Contador do per�odo
    linha = 0 # Contador de linha
    # Insere o n�mero do per�odo na coluna criada
    # Tira os valores desnecess�rios das linhas
    nome_coluna = [] # Cria uma lista contendo o nome de cada coluna
    for c in df_aulas.columns:
        nome_coluna.append(c)
    for dado in df_aulas['Ordem']:
        if (dado == f'{cont+1}� PER�ODO'):
            cont+=1
        if (str(dado)!='nan'):
            df_aulas['Per�odo'][linha] = str(int(cont))
            if (str(dado)=='Ordem' or str(dado)=='TOTAL' or str(dado)==f'{cont}� PER�ODO'):
                df_aulas['Per�odo'][linha] = nan
                coluna = 0
                while coluna < 11:
                    df_aulas[nome_coluna[coluna]][linha] = nan
                    coluna+=1
        linha+=1
    return df_aulas

# Verifica quais colunas possuem dados a serem validados
# Apaga as colunas desnecess�rias/n�o nomeadas
def ajusta_colunas(df):
    colunas = []
    drop_colunas = []
    contador = 0
    for c in df.head():
        # Feito desse modo para assegurar que a coluna ser� encontrada
        if (c != f'Unnamed: {contador-1}' and 
            c != f'Unnamed: {contador}' and 
            c != f'Unnamed: {contador+1}'):
            colunas.append(str(c))
        else:
            drop_colunas.append(str(c))
        contador+=1
    for dc in range(len(drop_colunas)):
        df.drop(drop_colunas[dc],inplace=True,axis=1)
    return df

# Insere os dados conforme o template
def insere_dados(df_para, df_aulas):
    for col_aulas in df_aulas.head():
        for col_para in df_para.head():
            dados_col_aula = []
            cont_dados_aulas = 0
            if (' '+col_aulas == col_para or
                col_aulas == col_para or
                col_aulas+' ' == col_para or
                (col_aulas == 'CRED' and col_para == 'Cr�dito')):
                for dados_aula in df_aulas[col_aulas]:
                    dados_col_aula.append(str(dados_aula))
                    cont_dados_aulas+=1
                for i in range(len(dados_col_aula)):
                    try:
                        # Se o dado inserido puder ser transformado em int, faz a convers�o
                        df_para[col_para][i] = int(dados_col_aula[i].split('.')[0])
                    except:
                        # Caso contr�rio, � armazenado como String
                        df_para[col_para][i] = dados_col_aula[i]
    return df_para

def converte_dtype(df_para):
    for col_para in df_para.head():
        dado_para = []
        if (col_para == 'AT' or
            col_para == 'AP' or
            col_para == 'TOTAL' or
            col_para == 'Cr�dito' or
            col_para == 'HA' or
            col_para == 'HR' or
            col_para == 'HR EAD' or
            col_para == 'CH (HR) Extensionista' or
            col_para == 'N�mero de Vagas' or
            col_para == 'Vagas Totais' or
            col_para == 'MT' or
            col_para == 'MP' or
            col_para == 'CH Docente Te�rica' or
            col_para == 'CH Docente Pr�tica' or
            col_para == 'CH Docente Total'):
            for dp in df_para[col_para]:
                dado_para.append(dp)
            for t_dp in range(len(dado_para)):
                try:
                    df_para[col_para][t_dp] = int(dado_para[t_dp])
                except:
                    continue
    return df_para

# Atentar a forma como a fun��o � escrita: Deve ser feita em ingl�s
def ajusta_funcoes(df_para):
    inicio = 14
    fim = 6
    for col_para in df_para.head():
        # Faltam algumas colunas para colocar
        if col_para == 'TOTAL':
            for i in range(len(df_para[col_para])):
                if ((len(df_para[col_para])-i) - fim > 0):
                    df_para[col_para][i] = f'=SUM(J{i+inicio}:K{i+inicio})'
                if ((((len(df_para[col_para])-i) - fim)) == 0):
                    df_para[col_para][i] = f'=SUM(L{inicio}:L{i+(inicio-1)})'
        elif col_para == 'Cr�dito':
            for i in range(len(df_para[col_para])):
                if ((len(df_para[col_para])-i) - fim > 0):
                    df_para[col_para][i] = f'=L{i+inicio}'
                if ((((len(df_para[col_para])-i) - fim)) == 0):
                    df_para[col_para][i] = f'=SUM(M{inicio}:M{i+(inicio-1)})'
        elif col_para == 'HA':
            for i in range(len(df_para[col_para])):
                if ((len(df_para[col_para])-i) - fim > 0):
                    df_para[col_para][i] = f'=L{i+inicio}*20' # Arrumar essa valida��o pq tem mais info para validar tem q usar o if
                if ((((len(df_para[col_para])-i) - fim)) == 0):
                    df_para[col_para][i] = f'=SUM(N{inicio}:N{i+(inicio-1)})'
        elif col_para == 'HR':
            for i in range(len(df_para[col_para])):
                if ((len(df_para[col_para])-i) - fim > 0):
                    df_para[col_para][i] = f'=L{i+inicio}*15'
                if ((((len(df_para[col_para])-i) - fim)) == 0):
                    df_para[col_para][i] = f'=SUM(O{inicio}:O{i+(inicio-1)})'
        elif col_para == 'HR EAD': # TODO: Essa valida��o de dados deu problema
            for i in range(len(df_para[col_para])):
        #        if ((len(df_para[col_para])-i) - fim > 0):
        #            df_para[col_para][i] = f'=IF(X{i+inicio}="EAD";M{i+inicio};0)'
                if ((((len(df_para[col_para])-i) - fim)) == 0):
                    df_para[col_para][i] = f'=SUM(P{inicio}:P{i+(inicio-1)})'
        elif col_para == 'CH (HR) Extensionista': # TODO: Falta a refer�ncia correta
            for i in range(len(df_para[col_para])):
        #        if ((len(df_para[col_para])-i) - fim > 0):
        #            df_para[col_para][i] = f'=SUM(H{i+inicio}:I{i+inicio})'
                if ((((len(df_para[col_para])-i) - fim)) == 0):
                    df_para[col_para][i] = f'=SUM(Q{inicio}:Q{i+(inicio-1)})'
        elif col_para == 'CH Docente Te�rica': # TODO: Tem uma valida��o imensa
            for i in range(len(df_para[col_para])):
        #        if ((len(df_para[col_para])-i) - fim > 0):
        #            df_para[col_para][i] = f'=SUM(H{i+inicio}:I{i+inicio})'
                if ((((len(df_para[col_para])-i) - fim)) == 0):
                    df_para[col_para][i] = f'=SUM(AC{inicio}:AC{i+(inicio-1)})'
        elif col_para == 'CH Docente Pr�tica': # TODO: Tem uma valida��o imensa
        #    for i in range(len(df_para[col_para])):
        #        if ((len(df_para[col_para])-i) - fim > 0):
        #            df_para[col_para][i] = f'=SUM(H{i+inicio}:I{i+inicio})'
                if ((((len(df_para[col_para])-i) - fim)) == 0):
                    df_para[col_para][i] = f'=SUM(AD{inicio}:AD{i+(inicio-1)})' # TODO: Por algum motivo n�o est� funcionando
        elif col_para == 'CH Docente Total':
            for i in range(len(df_para[col_para])):
                if ((len(df_para[col_para])-i) - fim > 0):
                    df_para[col_para][i] = f'=SUM(AC{i+inicio}:AD{i+inicio})'
                if ((((len(df_para[col_para])-i) - fim)) == 0):
                    df_para[col_para][i] = f'=SUM(AE{inicio}:AE{i+(inicio-1)})'
    return df_para

# Fun��o principal
def main():
    # Sele��o de arquivo dos aulas
    print("Selecione a planilha de aulas")
    Tk().withdraw()
    planilha_aulas = askopenfilename(
        filetypes=[('Arquivo excel', '.xlsx')],
        title='Selecione a planilha de aulas')
    # Mostra qual foi o arquivo selecionado das aulas
    print("Arquivo selecionado: " + str(planilha_aulas).split('/')[-1])
    
    # Selecionando a planilha PARA
    print("Selecione a planilha PARA")
    Tk().withdraw()
    planilha_para = askopenfilename(
        filetypes=[('Arquivo excel', '.xlsx')],
        title='Selecione a planilha PARA')
    # Mostra qual foi o arquivo selecionado PARA
    print("Arquivo selecionado: " + str(planilha_para).split('/')[-1])

    ''' -------------------------------------------------------------------------------
        Leitura das p�ginas do arquivo
        ------------------------------------------------------------------------------- '''
    
    print("Lendo arquivo")
    
    # Planilha de aulas � armazenado na vari�vel
    df_aulas = pd.read_excel(planilha_aulas, dtype=str, header=1)

    # planilha PARA e armazenada em uma variavel, selecionando apartir do header 12
    df_para = pd.read_excel(planilha_para, dtype=str, header=12, index_col=1, sheet_name=0)

    df_para = ajusta_colunas(df_para) # Faz os ajustes necess�rios nas colunas

    df_aulas.insert(loc=0, column='Per�odo', value=nan) # Cria uma coluna para colocar o n�mero do per�odo
    df_aulas = ajusta_info(df_aulas) # Ajusta as informa��es e linhas da tabela de aulas
    
    # Remove os nan das c�lulas vazias
    df_aulas.dropna(inplace=True, how='all') # remove a linha, quando todas as c�lulas estiverem vazias
    df_aulas.drop_duplicates(inplace=True) # remove duplicadas, no caso, os espa�os em branco e os t�tulos
    df_aulas = df_aulas.fillna("")
    
    #df_para.dropna(inplace=True, how='all') # remove a linha, quando todas as c�lulas estiverem vazias
    df_para.drop_duplicates(inplace=True) # remove duplicadas, no caso, os espa�os em branco e os t�tulos
    df_para = df_para.fillna("")
    
    df_para = insere_dados(df_para, df_aulas) # Pega os dados do dataframe de alunos e insere no template
    df_para = ajusta_funcoes(df_para) # TODO: Inserir as fun��es no local correto
    df_para = converte_dtype(df_para)
    
    df_para.to_excel('template_matriz.xlsx', index=False, startrow=12, startcol=2) # Cria a planilha utilizando o template PARA
    '''df_aulas.to_excel('matriz_aulas.xlsx', index=False) # Cria planilha do template DE'''

if __name__ == '__main__':
    # Chamada da fun��o main
    main()
