# -*- coding: Windows-1252 -*-

from tkinter import Tk
from tkinter.filedialog import askopenfilename
from numpy import nan
import pandas as pd

def ajusta_info(df_aulas):
    cont = 1 # Contador do período
    linha = 0 # Contador de linha
    # Insere o número do período na coluna criada
    # Tira os valores desnecessários das linhas
    nome_coluna = [] # Cria uma lista contendo o nome de cada coluna
    for c in df_aulas.columns:
        nome_coluna.append(c)
    for dado in df_aulas['Ordem']:
        if (dado == f'{cont+1}º PERÍODO'):
            cont+=1
        if (str(dado)!='nan'):
            df_aulas['Período'][linha] = str(int(cont))
            if (str(dado)=='Ordem' or str(dado)=='TOTAL' or str(dado)==f'{cont}º PERÍODO'):
                df_aulas['Período'][linha] = nan
                coluna = 0
                while coluna < 11:
                    df_aulas[nome_coluna[coluna]][linha] = nan
                    coluna+=1
        linha+=1
    return df_aulas

# Verifica quais colunas possuem dados a serem validados
# Apaga as colunas desnecessárias/não nomeadas
def ajusta_colunas(df):
    colunas = []
    drop_colunas = []
    contador = 0
    for c in df.head():
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
                (col_aulas == 'CRED' and col_para == 'Crédito')):
                for dados_aula in df_aulas[col_aulas]:
                    dados_col_aula.append(str(dados_aula))
                    cont_dados_aulas+=1
                for i in range(len(dados_col_aula)):
                    try:
                        # Se o dado inserido puder for int, faz a conversão
                        df_para[col_para][i] = int(dados_col_aula[i].split('.')[0])
                    except:
                        # Caso contrário, é armazenado como String
                        df_para[col_para][i] = dados_col_aula[i]
    return df_para

# Deveria permitir que as funções fossem criadas, porém ocorre o erro #NOME?
'''def ajusta_funcoes(df_para):
    inicio = 2
    for col_para in df_para.head():
        if col_para == 'TOTAL':
            for i in range(len(df_para[col_para])):
                df_para[col_para][i] = f'=SOMA(H{i+inicio}:I{i+inicio})'
    return df_para'''

# Função principal
def main():
    # Seleção de arquivo dos aulas
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
        Leitura das páginas do arquivo
        ------------------------------------------------------------------------------- '''
    
    print("Lendo arquivo")
    
    # Planilha de aulas é armazenado na variável
    df_aulas = pd.read_excel(planilha_aulas, dtype=str, header=1)

    # planilha PARA e armazenada em uma variavel, selecionando apartir do header 12
    df_para = pd.read_excel(planilha_para, dtype=str, header=12, index_col=1, sheet_name=0)

    df_para = ajusta_colunas(df_para) # Faz os ajustes necessários nas colunas

    df_aulas.insert(loc=0, column='Período', value=nan) # Cria uma coluna para colocar o número do período
    df_aulas = ajusta_info(df_aulas) # Ajusta as informações e linhas da tabela de aulas
    
    # Remove os nan das células vazias
    df_aulas.dropna(inplace=True, how='all') # remove a linha, quando todas as células estiverem vazias
    df_aulas.drop_duplicates(inplace=True) # remove duplicadas, no caso, os espaços em branco e os títulos
    df_aulas = df_aulas.fillna("")
    
    #df_para.dropna(inplace=True, how='all') # remove a linha, quando todas as células estiverem vazias
    df_para.drop_duplicates(inplace=True) # remove duplicadas, no caso, os espaços em branco e os títulos
    df_para = df_para.fillna("")

    df_para = insere_dados(df_para, df_aulas) # Pega os dados do dataframe de alunos e insere no template
    '''df_para = ajusta_funcoes(df_para) # TODO: Inserir as funções no local correto'''

    df_para.to_excel('template_matriz.xlsx', index=False) # Cria a planilha utilizando o template PARA
    '''df_aulas.to_excel('matriz_aulas.xlsx', index=False) # Cria planilha do template DE'''

if __name__ == '__main__':
    # Chamada da função main
    main()
