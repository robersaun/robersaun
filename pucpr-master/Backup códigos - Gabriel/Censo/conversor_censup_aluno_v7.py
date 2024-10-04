# -*- coding: Windows-1252 -*-

from tkinter import Tk
from tkinter.filedialog import askopenfilename
import pandas as pd

# Cod.: A04 -- Fun��o para verificar a validade e a exist�ncia do CPF
def trata_cpf(cpf):
    # Verifica se o CPF foi encontrado
    if cpf == "":
        return "CPF n�o encontrado na base de dados"
    resultado = ""
    for digito in cpf:
        if digito in '0123456789':
            resultado += digito
    while len(resultado) < 11:
        resultado = "0" + resultado
    # Verifica se o CPF possui apenas digitos identicos
    # Se for tudo identico - o CPF � invalido
    if (resultado == "00000000000" or
        resultado == "11111111111" or
        resultado == "22222222222" or
        resultado == "33333333333" or
        resultado == "44444444444" or
        resultado == "55555555555" or
        resultado == "66666666666" or
        resultado == "77777777777" or
        resultado == "88888888888" or
        resultado == "99999999999"):
        return "Preencher com um CPF v�lido"
    mVal1 = 10
    cont1 = 0
    rVal1 = 0
    while mVal1 > 1:
        num = int(resultado[cont1])
        rVal1 = (num * mVal1) + rVal1
        mVal1 -= 1
        cont1 += 1
    verVal1 = str((rVal1 * 10)%11)
    # Verifica se o resultado da soma e divis�o dos 9 primeiros d�gitos � igual ao d�gito 10
    # Se o resultado for diferente do 10� d�gito - o CPF � invalido
    if verVal1[-1] != resultado[cont1]:
        return "Preencher com um CPF v�lido"
    mVal2 = 11
    cont2 = 0
    rVal2 = 0
    while mVal2 > 1:
        num = int(resultado[cont2])
        rVal2 = (num * mVal2) + rVal2
        mVal2 -= 1
        cont2 += 1
    verVal2 = str((rVal2 * 10)%11)
    # Verifica se o resultado da soma e divis�o dos 10 primeiros d�gitos � igual ao d�gito 11
    # Se o resultado for diferente do 11� d�gito - o CPF � invalido
    if verVal2[-1] != resultado[cont2]:
        return "Preencher com um CPF v�lido"
    # Se passar por todas as verifica��es de valida��o retorna o CPF
    else:
        return resultado

# Cod.: A06.1 -- Insere as barras nas datas para poder ocorrer a transofmra��o em date time
def ajusta_formato(data):
    if len(data) <= 10 and len(data) > 0:
        cont = 0
        resultado = ""
        resultado_dt = ""
        data = str(data)
        for digito in data:
            if digito in '0123456789':
                resultado += digito
        while len(resultado) < 8:
            resultado = "0" + resultado
        for digito in resultado:
            if cont == 2 or cont == 4:
                resultado_dt += '/'
            if digito in '0123456789':
                resultado_dt += digito
            cont += 1
        return resultado_dt
    else:
        return data

# Cod.: A06.2 -- Remove as barras da data
def trata_data(data):
    resultado = ""
    data = str(data)
    for digito in data:
        if digito in '0123456789':
            resultado += digito
    while len(resultado) < 8:
        resultado = "0" + resultado
    return resultado

# Cod.: GED -- Para cada discente, escreve os dados pessoais
def escreve_txt(df_alunos, df_vinculos, dados):
    for index, linha in df_alunos.iterrows():
        escrita = "|".join(linha)
        dados.write(escrita + "\n")

        # Busca no dataframe de v�nculo todos os registros dos alunos (index)
        df_comparacao = df_vinculos.loc[df_vinculos['CPF'].eq(linha[3])]

        # Para cada v�nculo, escreve os dados de v�nculo
        for ind, comparacao in df_comparacao.iterrows():
            verificaCPF = "".join(comparacao[0:1])
            # Verifica se h� algum CPF para criar o vinculo
            if (verificaCPF != 'Preencher com um CPF v�lido' and 
                verificaCPF != 'CPF n�o encontrado na base de dados'):
                escrita = "|".join(comparacao[1:])
                dados.write(escrita + "\n")

# Fun��o principal
def main():
    # Sele��o de arquivo dos alunos
    print("Selecione a planilha de alunos")
    Tk().withdraw()
    planilha_alunos = askopenfilename(
        filetypes=[('Arquivo excel', '.xlsx')],
        title='Selecione a planilha de alunos')

    # Mostra qual foi o arquivo selecionado
    print("Arquivo selecionado: " + str(planilha_alunos).split('/')[-1])

    ''' -------------------------------------------------------------------------------
        Leitura das p�ginas do arquivo
        ------------------------------------------------------------------------------- '''
    
    print("Lendo arquivo")

    # Planilha de alunos � armazenado na vari�vel
    # Os dados dos alunos dever�o sempre estar na primeira p�gina do arquivo
    df_alunos = pd.read_excel(planilha_alunos, dtype=str, sheet_name=0)
    # Pega as informa��es da p�gina de v�nculos
    # Os dados dos vinculos dever�o sempre estar na segunda p�gina do arquivo
    df_vinculos = pd.read_excel(planilha_alunos, dtype=str, sheet_name=1)

    # Remove os nan das c�lulas vazias
    df_alunos = df_alunos.fillna("")
    df_vinculos = df_vinculos.fillna("")

    ''' --------------------------------------------------------------------------------
        Leitura das informa��es e tratamento dos dados do arquivo
        -------------------------------------------------------------------------------- '''
    
    print("Identificando colunas")

    # Ocorre a identifica��o das colunas das p�ginas sobre alunos 
    # Para os casos em que h� chance de altera��o no nome da referida coluna
    colunas_aluno = []
    for nome_coluna_a in range(len(df_alunos.columns)):
        colunas_aluno.append(df_alunos.columns[nome_coluna_a])
    # Ocorre o mesmo, mas para os v�nculos
    colunas_vinculo = []
    for nome_coluna_v in range(len(df_vinculos.columns)):
        colunas_vinculo.append(df_vinculos.columns[nome_coluna_v])

    print("Tratando dados")
    print("- Tratando dados de CPF")

    # Cod.: A04 -- Chama a fun��o de tratar os CPFs
    df_alunos['CPF'] = df_alunos['CPF'].apply(lambda cpf: trata_cpf(cpf))
    df_vinculos['CPF'] = df_vinculos['CPF'].apply(lambda cpf: trata_cpf(cpf))

    print("> Tratamento de CPF conclu�do")
    print("- Tratando dados de Data")

    # Cod.: A06 -- Converte a coluna Data Nascimento para o formato de data
    # Importante: A coluna contendo a informa��o sobre a data de nascimento do aluno n�o dever� mudar de lugar 
    # Caso a altera��o ocorra, mudar o a posi��o buscada
    data_nascimento = colunas_aluno[5] 
    df_alunos[data_nascimento] = df_alunos[data_nascimento].apply(lambda data: ajusta_formato(data))
    df_alunos[data_nascimento] = pd.to_datetime(df_alunos[data_nascimento], errors='coerce')
    df_alunos[data_nascimento] = df_alunos[data_nascimento].dt.strftime('%d/%m/%Y')
    df_alunos[data_nascimento] = df_alunos[data_nascimento].apply(lambda data: trata_data(data))

    print("> Tratamento de Data conclu�do")
    print("- Removendo espa�os vazios")

    # Remove os espa�os do in�cio e final das c�lulas, caso haja algum
    df_alunos[df_alunos.columns] = df_alunos.apply(lambda x: x.str.strip())
    df_vinculos[df_vinculos.columns] = df_vinculos.apply(lambda x: x.str.strip())

    print("> Remo��o conclu�da")
    ''' --------------------------------------------------------------------------------
        Escrita das informa��es no .txt
        -------------------------------------------------------------------------------- '''

    print("Escrevendo .txt")

    # Cria o arquivo dados e escreve a primeira linha com 40|10
    # Enconding em Windows-1252 para escrever corretamente a valida��o do CPF
    dados = open("dados_alunos.txt", "w+", encoding='Windows-1252')
    dados.write("40|10\n")

    # Cod.: GED -- Escreve as informa��es no xt
    escreve_txt(df_alunos, df_vinculos, dados)

    print(".txt conclu�do")


if __name__ == '__main__':
    # Chamada da fun��o main
    main()
