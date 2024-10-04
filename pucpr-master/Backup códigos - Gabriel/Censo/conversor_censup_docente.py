from tkinter import Tk
from tkinter.filedialog import askopenfilename

import pandas as pd


# Função para tratar o CPF, retirando a pontuação ou caracteres especiais
# e deixando o número certo de digitos
def trata_cpf(cpf):
    resutado = ""

    for digito in cpf:
        if digito in '0123456789':
            resutado += digito

    while len(resutado) < 11:
        resutado = "0" + resutado

    return resutado


# Função para tratar a data, retirando a pontuação ou caracteres especiais
# e deixando o número certo de digitos
def trata_data(data):
    resutado = ""
    data = str(data)

    for digito in data:
        if digito in '0123456789':
            resutado += digito

    while len(resutado) < 8:
        resutado = "0" + resutado

    return resutado


# Função principal
def main():
    # Seleção de arquivo dos docentes
    print("Selecione a planilha de docentes")
    Tk().withdraw()
    planilha_docentes = askopenfilename(
        filetypes=[('Arquivo excel', '.xlsx')],
        title='Selecione a planilha de docentes')

    # Planilha de docentes é armazenado na variável df_docentes e planilha vínculo é armazenada no df_vinculo
    df_docentes = pd.read_excel(planilha_docentes, dtype=str, sheet_name="Dados Pessoais")
    df_vinculos = pd.read_excel(planilha_docentes, dtype=str, sheet_name="Vínculo")

    # Remove os nan das células vazias
    df_docentes = df_docentes.fillna("")

    # Chama a função de tratar CPF para cada CPF da tabela de docentes e na tabela de vínculo
    df_docentes['CPF'] = df_docentes['CPF'].apply(lambda cpf: trata_cpf(cpf))
    df_vinculos['CPF'] = df_vinculos['CPF'].apply(lambda cpf: trata_cpf(cpf))

    # Converte a coluna Data de Nascimento para o formato de data no df_docentes
    df_docentes['Data de Nascimento'] = pd.to_datetime(df_docentes['Data de Nascimento'], errors='coerce')
    df_docentes["Data de Nascimento"] = df_docentes["Data de Nascimento"].dt.strftime("%d/%m/%Y")
    df_docentes['Data de Nascimento'] = df_docentes['Data de Nascimento'].apply(lambda data: trata_data(data))

    # Remove os espaços do início e final das células, caso haja algum
    df_docentes[df_docentes.columns] = df_docentes.apply(lambda x: x.str.strip())
    df_vinculos[df_vinculos.columns] = df_vinculos.apply(lambda x: x.str.strip())

    # Cria o arquivo dados e escreve a primeira linha com 30|10 (padrão censo)
    dados = open("dados.txt", "w+", encoding='utf-8')
    dados.write("30|10\n")

    # Para cada docente, escreve os dados pessoais
    for index, linha in df_docentes.iterrows():
        escrita = "|".join(linha)
        dados.write(escrita + "\n")

        # Busca no dataframe de vínculo, todos os regístros do professor atual (index)
        df_comparacao = df_vinculos.loc[df_vinculos['CPF'].eq(linha[3])]

        # Para cada vínculo, escreve os dados de vínculo
        for ind, comparacao in df_comparacao.iterrows():
            escrita = "|".join(comparacao[1:])
            dados.write(escrita + "\n")


if __name__ == '__main__':
    # Chamada da função main
    main()
