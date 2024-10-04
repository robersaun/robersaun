# -*- coding: Windows-1252 -*-

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
def verifica_estrangeiro(docE, nacionalidade):
    nacionalidade = str(nacionalidade)
    docE = str(docE)
    # Assim ainda dá errado >> Fica td 0
    if nacionalidade == '':
        return "Nacionalidade não informada"
    elif nacionalidade == '1' or nacionalidade == '2':
        return "Preenchimento exclusivo para alunos estrangeiros"
    else:
        return docE

# verifica cor e raça do estudante
# TODO: Fazer a verificação do ano de ingresso
# TODO: Ver como quebrar a data
def verifica_cr(cr):
    if cr == "":
        return "Cor/Raça não informada"
    if cr in "0123456":
        return cr
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

# Verifica se o aluno possui alguma deficiência
def tem_necessidade(tn, c, bv, s, a, sc):
    tn = str(tn) # Indica se o aluno possui alguma necessidade
    # As letras abaixo indicam o tipo de necessidade:
    c = str(c) # Cegueira
    bv = str(bv) # Baixa Visão
    s = str(s) # Surdez
    a = str(a) # Auditiva
    sc = str(sc) #Surdo-cegueira
    # Escreve o texto com os avisos
    possui_necessidade = ''
    
    # Verificação se possui a necessidade
    if tn == '0':
        possui_necessidade += 'Não'
    elif tn == '1':
        possui_necessidade += 'Sim'
    elif tn == '2':
        possui_necessidade += 'Não dispõe da informação'
    else:
        possui_necessidade += 'Prencher com um dos valores: 0, 1, 2'
    
    # Verificação se possui marcado mais de uma necessidade relacionada a visão
    if (((c == '1' and bv == '1') or 
        (c == '1' and sc == '1') or 
        (bv == '1' and sc == '1') or 
        (c == '1' and bv == '1' and sc == '1')) and 
        tn == '1'):
        possui_necessidade += ' Não é possível informar Cegueira ou Baixa Visão ou Surdocegueira simultaneamente'
    elif (((c == '1' and bv == '1') or
         (c == '1' and sc == '1') or 
         (bv == '1' and sc == '1') or 
         (c == '1' and bv == '1' and sc == '1')) and 
         (tn == '0' or tn=='2')):
        possui_necessidade += ' Não pode ser alterado. O aluno está vinculado a um programa de reserva de vagas para deficiêntes em outra IES.'

    # Verificação se possui marcado mais de uma necessidade relacionada a audiçã
    if (((s == '1' and a == '1') or 
        (s == '1' and sc == '1') or 
        (a == '1' and sc == '1') or 
        (s == '1' and a == '1' and sc == '1')) and 
        tn == '1'):
        possui_necessidade += ' Não é possível informar Surdez ou Deficiêmcia Auditiva ou Surdocegueira simultaneamente'
    elif (((s == '1' and a == '1') or 
          (s == '1' and sc == '1') or 
          (a == '1' and sc == '1') or 
          (s == '1' and a == '1' and sc == '1')) and 
          (tn == '0' or tn=='2')):
        possui_necessidade += ' Não pode ser alterado. O aluno está vinculado a um programa de reserva de vagas para deficiêntes em outra IES.'

    # Verificação se possui marcado mais de uma necessidade relacionada a visão e audição
    if (((c == '1' and s == '1') or 
        (c == '1' and sc == '1') or 
        (s == '1' and sc == '1') or 
        (c == '1' and s == '1' and sc == '1')) and 
        tn == '1'):
        possui_necessidade += ' Não é possível informar Cegueira ou Surdez ou Surdocegueira simultaneamente'
    elif (((c == '1' and s == '1') or 
          (c == '1' and sc == '1') or 
          (s == '1' and sc == '1') or 
          (c == '1' and s == '1' and sc == '1')) and 
          (tn == '0' or tn=='2')):
        possui_necessidade += ' Não pode ser alterado. O aluno está vinculado a um programa de reserva de vagas para deficiêntes em outra IES.'

    return possui_necessidade

# Verifica se o campo da deficiência foi preenchido
def valida_campo_deficiencia(tn, d):
    tn = str(tn) # Idicativo se tem ou não alguma deficiência
    d = str(d) # Nome da deficiência
    if (tn == 'Sim' and d == ''):
        return 'Preenchimento obrigatório para alunos com deficiência'
    elif (tn == 'Sim' and d == '0'):
        return 'Não'
    elif (tn == 'Sim' and d == '1'):
        return 'Sim'
    elif (tn == 'Sim' and (d!='' or d!= '0' or d!='1')):
        return 'Preencher com um dos valores: (vazio), 0, 1'
    elif (tn == 'Não' or tn == 'Não dispõe da informação'):
        return 'Preenchimento exclusivo para alunos com deficiência'
    else:
        # O que deveria ocorrer se nada tiver sido informado em 'Tem Necessidade'?
        return '-1'

#Verifica qual foi o turno selecionado
def verifica_turno(turno):
    if turno == "1":
        return "Matutino"
    elif turno == "2":
        return "Vespertino"
    elif turno == "3":
        return "Noturno"
    elif turno == "4":
        return "Integral"
    else:
        return "Preencher com um dos valores: 1, 2, 3, 4"

# Para cada discente, escreve os dados pessoais
def escreve_txt(escola, dados):
    for index, linha in escola.iterrows():
        escrita = "|".join(linha)
        # TODO: Ver oq fazer com isso > Por conta do vínculo com o curso
        dados.write("41|" + escrita + "\n")


# Função principal
def main():
    # Seleção de arquivo dos alunos
    print("Selecione a planilha de alunos")
    Tk().withdraw()
    planilha_alunos = askopenfilename(
        filetypes=[('Arquivo excel', '.xlsx')],
        title='Selecione a planilha de alunos')

    # Planilha de alunos é armazenado na variável conforme a escola 
    # Sendo respectivamente: escola de direito, belas artes, educação e humanidade ou medicina
    df_direito = pd.read_excel(planilha_alunos, dtype=str, sheet_name="Direito")
    df_belas = pd.read_excel(planilha_alunos, dtype=str, sheet_name="Belas")
    df_eh = pd.read_excel(planilha_alunos, dtype=str, sheet_name="EH")
    df_medicina = pd.read_excel(planilha_alunos, dtype=str, sheet_name="Medicina")

    # A princípio tem que tratar as linhas que só te as informações de totais
    # Talvez influencie na questão do curso
    # Concatena as páginas, em uma única
    df_alunos = pd.concat([df_direito, df_belas, df_eh, df_medicina])

    # Remove os nan das células vazias
    df_alunos = df_alunos.fillna("")

    # Chama a função de tratar CPF para cada CPF de cada tabela de alunos
    df_alunos['CPF'] = df_alunos['CPF'].apply(lambda cpf: trata_cpf(cpf))

    # Verifica se há documento de estrangeiro
    df_alunos['Documento Estrangeiro'] = df_alunos.apply(lambda n:
        verifica_estrangeiro(n['Documento Estrangeiro'], n['Nacionalidade']), axis=1)

    # Verifica cor e raça
    df_alunos['Raça'] = df_alunos['Raça'].apply(lambda cr: verifica_cr(cr))

    # Verifica Turno
    df_alunos['Turno'] = df_alunos['Turno'].apply(lambda turno: verifica_turno(turno))

    # Verifica local de conclusão do Ensino Médio
    df_alunos['Tipo Escola'] = df_alunos['Tipo Escola'].apply(lambda em: verifica_em(em))

    # Verificar alunos com deficiência
    df_alunos['Tem Necessidade'] = df_alunos.apply(lambda tn: 
        tem_necessidade(tn['Tem Necessidade'], # genérico
                        tn['Cegueira'], tn['Baixa Visão'], # relacionados a visão
                        tn['Surdez'], tn['Auditiva'], # relacionado a audição
                        tn['Surdo-cegueira'] # audição e visão
                        ), axis=1)

    deficiencia = ['Cegueira', 'Baixa Visão', 
                    'Surdez', 'Auditiva', 
                    'Física', 
                    'Surdo-cegueira', 
                    'Múltipla', 
                    'Mental', 'Autismo infantil', 'Síndrome de Asperger', 'Síndrome de Rett', 'Transtorno Desintegrativo da Infância', 
                    'Altas habilidades/superdotação']
    for nome_deficiencia in deficiencia:
        df_alunos[nome_deficiencia] = df_alunos.apply(lambda d: 
            valida_campo_deficiencia(d['Tem Necessidade'], d[nome_deficiencia]), axis=1)

    # Remove os espaços do início e final das células, caso haja algum
    df_alunos[df_alunos.columns] = df_alunos.apply(lambda x: x.str.strip())

    # Cria o arquivo dados e escreve a primeira linha com 40|10
    dados = open("dados_alunos.txt", "w+", encoding='utf-8')
    # TODO: Ver se isso se manterá assim
    dados.write("40|10\n")

    # Escreve as informações no xt
    escreve_txt(df_alunos, dados)

if __name__ == '__main__':
    # Chamada da função main
    main()

# Cabeçalho do arquivo - Registro 40: A princípio OK
# Registro do aluno - Registro 41: Não OK
# Ok > 2, 4 , 5, 12 - 21
# Meio OK > 7 (Separar a string do ano)
# Problema de com 2 colunas: 8 (Precisa de outro doc), 9 (Precisa de outro doc), 10 , 11
# ?: 11
# Dúvidas > (3) Que base? Deveríamos comparar com outra?
# Registro do vínculo do aluno com o curso - Registro 42: Não OK
# Ok > 6

# Corrigir as linhas que informam o Total de Registro e outras que não informam dados de alunos