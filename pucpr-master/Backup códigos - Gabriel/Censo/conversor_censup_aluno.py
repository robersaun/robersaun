from tkinter import Tk
from tkinter.filedialog import askopenfilename
import pandas as pd

# Função para tratar o CPF, retirando a pontuação ou caracteres especiais
# e deixando o número certo de digitos
def trata_cpf(cpf):
    resultado = ""
    verificacao = 0
    for digito in cpf:
        if digito in '0123456789':
            resultado += digito
            verificacao += int(digito)
    while len(resultado) < 11:
        resultado = "0" + resultado
    verificacao = str(verificacao)
    # Verifica se o CPF foi encontrado
    if resultado == "00000000000":
        return "CPF não encontrado na base de dados"
    # Verifica se o CPF contem é válido
    elif verificacao[0] != verificacao[-1]:
        return "Preencher com um CPF válido"
    else:
        return resultado

# Verifica se o aluno é estrangeiro ou não
# TODO: Verificar como retornar o número da documentação
def verifica_estrangeiro(nacionalidade):
    nacionalidade = str(nacionalidade)
    if nacionalidade == "":
        return "Nacionalidade não informada"
    elif nacionalidade == "1" or nacionalidade == "2":
        return "Preenchimento exclusivo para alunos estrangeiros"
    else:
        return "Verificar mais tarde"

# verifica qual a raça do estudante
# TODO: Fazer a verificação do ano de ingresso
def verifica_raca(raca):
    if raca == "":
        return "Cor/Raça não informada"
    if raca in "0123456":
        return raca
    else:
        return "Preencher com um dos valore: 0, 1, 2, 3, 4, 5, 6"

# Verifica qual o foi o tipo de Ensino Médio que o estudante se formou
def verifica_em(em):
    if em == "0":
        return "Privado"
    elif em == "1":
        return "Publico"
    elif em == "2":
        return "Não dispõe da informação"
    else:
        return "Preencher com um dos valores: 0, 1, 2"


# Para cada docente, escreve os dados pessoais
def escreve_txt(escola, dados):
    for index, linha in escola.iterrows():
        escrita = "|".join(linha)
        # TODO: Ver oq fazer com isso > Por conta do vínculo com o curso
        # dados.write("41|" + escrita + "\n")
        dados.write(escrita + "\n")

# Função principal
def main():
    # Seleção de arquivo dos alunos
    print("Selecione a planilha de alunos")
    Tk().withdraw()
    planilha_alunos = askopenfilename(
        filetypes=[('Arquivo excel', '.xlsx')],
        title='Selecione a planilha de alunos')

    # Planilha de alunos é armazenado na variável conforme a escola 
    # df_direito para escola de direito
    # df_belas para escola de belas artes
    # df_eh para escola de educação e humanidade
    # df_medicina para escola de medicina
    # >> df_alunos = pd.read_excel(planilha_alunos, dtype=str, sheet_name=["Direito", "Belas", "EH", "Medicina"])
    df_direito = pd.read_excel(planilha_alunos, dtype=str, sheet_name="Direito")
    df_belas = pd.read_excel(planilha_alunos, dtype=str, sheet_name="Belas")
    df_eh = pd.read_excel(planilha_alunos, dtype=str, sheet_name="EH")
    df_medicina = pd.read_excel(planilha_alunos, dtype=str, sheet_name="Medicina")

    # Remove os nan das células vazias
    # >> df_alunos = df_alunos.fillna("")
    df_direito = df_direito.fillna("")
    df_belas = df_belas.fillna("")
    df_eh = df_eh.fillna("")
    df_medicina = df_medicina.fillna("")

    # Chama a função de tratar CPF para cada CPF de cada tabela de alunos
    # >> df_alunos['CPF'] = df_alunos['CPF'].apply(lambda cpf: trata_cpf(cpf))
    df_direito['CPF'] = df_direito['CPF'].apply(lambda cpf: trata_cpf(cpf))
    df_belas['CPF'] = df_belas['CPF'].apply(lambda cpf: trata_cpf(cpf))
    df_eh['CPF'] = df_eh['CPF'].apply(lambda cpf: trata_cpf(cpf))
    df_medicina['CPF'] = df_medicina['CPF'].apply(lambda cpf: trata_cpf(cpf))

    # Verifica se há documento de estrangeiro
    # >> df_alunos['Documento Estrangeiro'] = df_alunos['Nacionalidade'].apply(lambda nacionalidade: 
    # >>     verifica_estrangeiro(nacionalidade))
    df_direito['Documento Estrangeiro'] = df_direito['Nacionalidade'].apply(lambda nacionalidade: 
        verifica_estrangeiro(nacionalidade))
    df_belas['Documento Estrangeiro'] = df_belas['Nacionalidade'].apply(lambda nacionalidade: 
        verifica_estrangeiro(nacionalidade))
    df_eh['Documento Estrangeiro'] = df_eh['Nacionalidade'].apply(lambda nacionalidade: 
        verifica_estrangeiro(nacionalidade))
    df_medicina['Documento Estrangeiro'] = df_medicina['Nacionalidade'].apply(lambda nacionalidade: 
        verifica_estrangeiro(nacionalidade))

    # Verifica cor e raça
    df_direito['Raça'] = df_direito['Raça'].apply(lambda raca: verifica_raca(raca))
    df_belas['Raça'] = df_belas['Raça'].apply(lambda raca: verifica_raca(raca))
    df_eh['Raça'] = df_eh['Raça'].apply(lambda raca: verifica_raca(raca))
    df_medicina['Raça'] = df_medicina['Raça'].apply(lambda raca: verifica_raca(raca))

    # Verifica local de conclusão do Ensino Médio
    df_direito['Tipo Escola'] = df_direito['Tipo Escola'].apply(lambda em: verifica_em(em))
    df_belas['Tipo Escola'] = df_belas['Tipo Escola'].apply(lambda em: verifica_em(em))
    df_eh['Tipo Escola'] = df_eh['Tipo Escola'].apply(lambda em: verifica_em(em))
    df_medicina['Tipo Escola'] = df_medicina['Tipo Escola'].apply(lambda em: verifica_em(em))

    # Remove os espaços do início e final das células, caso haja algum
    # >> df_alunos[df_alunos.columns] = df_alunos.apply(lambda x: x.str.strip())
    df_direito[df_direito.columns] = df_direito.apply(lambda x: x.str.strip())
    df_belas[df_belas.columns] = df_belas.apply(lambda x: x.str.strip())
    df_eh[df_eh.columns] = df_eh.apply(lambda x: x.str.strip())
    df_medicina[df_medicina.columns] = df_medicina.apply(lambda x: x.str.strip())

    # Cria o arquivo dados e escreve a primeira linha com 40|10
    dados = open("dados_alunos.txt", "w+", encoding='utf-8')
    # TODO: Ver se isso se manterá assim
    dados.write("40|10\n")

    # Escreve as informações no xt
    # >> escreve_txt(df_alunos, dados)
    escreve_txt(df_direito, dados)
    escreve_txt(df_belas, dados)
    escreve_txt(df_eh, dados)
    escreve_txt(df_medicina, dados)

if __name__ == '__main__':
    # Chamada da função main
    main()

# Cabeçalho do arquivo - Registro 40: A princípio OK
# Registro do aluno - Registro 41: Não OK
# Ok > 2, 4 
# Meio OK > 7
# Problema de com 2 colunas: 5, 8, 9, 10 , 11 - 21
# ?: 11
# Dúvidas > (3) Que base? Deveríamos comparar com outra?
# Registro do vínculo do aluno com o curso - Registro 42: Não OK

# Corrigir as linhas que informam o Total de Registro e outras que não informam dados de alunos