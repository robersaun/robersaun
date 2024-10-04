# -*- coding: Windows-1252 -*-

from tkinter import Tk
from tkinter.filedialog import askopenfilename
import pandas as pd

# Fun��o para tratar o CPF, retirando a pontua��o ou caracteres especiais
# e deixando o n�mero certo de digitos
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
        return "CPF n�o encontrado na base de dados"
    # Verifica se o CPF contem � v�lido
    elif verificacao[0] != verificacao[-1]:
        return "Preencher com um CPF v�lido"
    else:
        return resultado

# Verifica se o aluno � estrangeiro ou n�o
def verifica_estrangeiro(docE, nacionalidade):
    nacionalidade = str(nacionalidade)
    docE = str(docE)
    # Assim ainda d� errado >> Fica td 0
    if nacionalidade == '':
        return "Nacionalidade n�o informada"
    elif nacionalidade == '1' or nacionalidade == '2':
        return "Preenchimento exclusivo para alunos estrangeiros"
    else:
        return docE

# verifica cor e ra�a do estudante
# TODO: Fazer a verifica��o do ano de ingresso
# TODO: Ver como quebrar a data
def verifica_cr(cr):
    if cr == "":
        return "Cor/Ra�a n�o informada"
    if cr in "0123456":
        return cr
    else:
        return "Preencher com um dos valore: 0, 1, 2, 3, 4, 5, 6"

# Verifica qual o foi o tipo de Ensino M�dio que o estudante se formou
def verifica_em(em):
    if em == "0":
        return "Privado"
    elif em == "1":
        return "Publico"
    elif em == "2":
        return "N�o disp�e da informa��o"
    else:
        return "Preencher com um dos valores: 0, 1, 2"

# Verifica se o aluno possui alguma defici�ncia
def tem_necessidade(tn, c, bv, s, a, sc):
    tn = str(tn) # Indica se o aluno possui alguma necessidade
    # As letras abaixo indicam o tipo de necessidade:
    c = str(c) # Cegueira
    bv = str(bv) # Baixa Vis�o
    s = str(s) # Surdez
    a = str(a) # Auditiva
    sc = str(sc) #Surdo-cegueira
    # Escreve o texto com os avisos
    possui_necessidade = ''
    
    # Verifica��o se possui a necessidade
    if tn == '0':
        possui_necessidade += 'N�o'
    elif tn == '1':
        possui_necessidade += 'Sim'
    elif tn == '2':
        possui_necessidade += 'N�o disp�e da informa��o'
    else:
        possui_necessidade += 'Prencher com um dos valores: 0, 1, 2'
    
    # Verifica��o se possui marcado mais de uma necessidade relacionada a vis�o
    if (((c == '1' and bv == '1') or 
        (c == '1' and sc == '1') or 
        (bv == '1' and sc == '1') or 
        (c == '1' and bv == '1' and sc == '1')) and 
        tn == '1'):
        possui_necessidade += ' N�o � poss�vel informar Cegueira ou Baixa Vis�o ou Surdocegueira simultaneamente'
    elif (((c == '1' and bv == '1') or
         (c == '1' and sc == '1') or 
         (bv == '1' and sc == '1') or 
         (c == '1' and bv == '1' and sc == '1')) and 
         (tn == '0' or tn=='2')):
        possui_necessidade += ' N�o pode ser alterado. O aluno est� vinculado a um programa de reserva de vagas para defici�ntes em outra IES.'

    # Verifica��o se possui marcado mais de uma necessidade relacionada a audi��
    if (((s == '1' and a == '1') or 
        (s == '1' and sc == '1') or 
        (a == '1' and sc == '1') or 
        (s == '1' and a == '1' and sc == '1')) and 
        tn == '1'):
        possui_necessidade += ' N�o � poss�vel informar Surdez ou Defici�mcia Auditiva ou Surdocegueira simultaneamente'
    elif (((s == '1' and a == '1') or 
          (s == '1' and sc == '1') or 
          (a == '1' and sc == '1') or 
          (s == '1' and a == '1' and sc == '1')) and 
          (tn == '0' or tn=='2')):
        possui_necessidade += ' N�o pode ser alterado. O aluno est� vinculado a um programa de reserva de vagas para defici�ntes em outra IES.'

    # Verifica��o se possui marcado mais de uma necessidade relacionada a vis�o e audi��o
    if (((c == '1' and s == '1') or 
        (c == '1' and sc == '1') or 
        (s == '1' and sc == '1') or 
        (c == '1' and s == '1' and sc == '1')) and 
        tn == '1'):
        possui_necessidade += ' N�o � poss�vel informar Cegueira ou Surdez ou Surdocegueira simultaneamente'
    elif (((c == '1' and s == '1') or 
          (c == '1' and sc == '1') or 
          (s == '1' and sc == '1') or 
          (c == '1' and s == '1' and sc == '1')) and 
          (tn == '0' or tn=='2')):
        possui_necessidade += ' N�o pode ser alterado. O aluno est� vinculado a um programa de reserva de vagas para defici�ntes em outra IES.'

    return possui_necessidade

# Verifica se o campo da defici�ncia foi preenchido
def valida_campo_deficiencia(tn, d):
    tn = str(tn) # Idicativo se tem ou n�o alguma defici�ncia
    d = str(d) # Nome da defici�ncia
    if (tn == 'Sim' and d == ''):
        return 'Preenchimento obrigat�rio para alunos com defici�ncia'
    elif (tn == 'Sim' and d == '0'):
        return 'N�o'
    elif (tn == 'Sim' and d == '1'):
        return 'Sim'
    elif (tn == 'Sim' and (d!='' or d!= '0' or d!='1')):
        return 'Preencher com um dos valores: (vazio), 0, 1'
    elif (tn == 'N�o' or tn == 'N�o disp�e da informa��o'):
        return 'Preenchimento exclusivo para alunos com defici�ncia'
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
        # TODO: Ver oq fazer com isso > Por conta do v�nculo com o curso
        dados.write("41|" + escrita + "\n")


# Fun��o principal
def main():
    # Sele��o de arquivo dos alunos
    print("Selecione a planilha de alunos")
    Tk().withdraw()
    planilha_alunos = askopenfilename(
        filetypes=[('Arquivo excel', '.xlsx')],
        title='Selecione a planilha de alunos')

    # Planilha de alunos � armazenado na vari�vel conforme a escola 
    # Sendo respectivamente: escola de direito, belas artes, educa��o e humanidade ou medicina
    df_direito = pd.read_excel(planilha_alunos, dtype=str, sheet_name="Direito")
    df_belas = pd.read_excel(planilha_alunos, dtype=str, sheet_name="Belas")
    df_eh = pd.read_excel(planilha_alunos, dtype=str, sheet_name="EH")
    df_medicina = pd.read_excel(planilha_alunos, dtype=str, sheet_name="Medicina")

    # A princ�pio tem que tratar as linhas que s� te as informa��es de totais
    # Talvez influencie na quest�o do curso
    # Concatena as p�ginas, em uma �nica
    df_alunos = pd.concat([df_direito, df_belas, df_eh, df_medicina])

    # Remove os nan das c�lulas vazias
    df_alunos = df_alunos.fillna("")

    # Chama a fun��o de tratar CPF para cada CPF de cada tabela de alunos
    df_alunos['CPF'] = df_alunos['CPF'].apply(lambda cpf: trata_cpf(cpf))

    # Verifica se h� documento de estrangeiro
    df_alunos['Documento Estrangeiro'] = df_alunos.apply(lambda n:
        verifica_estrangeiro(n['Documento Estrangeiro'], n['Nacionalidade']), axis=1)

    # Verifica cor e ra�a
    df_alunos['Ra�a'] = df_alunos['Ra�a'].apply(lambda cr: verifica_cr(cr))

    # Verifica Turno
    df_alunos['Turno'] = df_alunos['Turno'].apply(lambda turno: verifica_turno(turno))

    # Verifica local de conclus�o do Ensino M�dio
    df_alunos['Tipo Escola'] = df_alunos['Tipo Escola'].apply(lambda em: verifica_em(em))

    # Verificar alunos com defici�ncia
    df_alunos['Tem Necessidade'] = df_alunos.apply(lambda tn: 
        tem_necessidade(tn['Tem Necessidade'], # gen�rico
                        tn['Cegueira'], tn['Baixa Vis�o'], # relacionados a vis�o
                        tn['Surdez'], tn['Auditiva'], # relacionado a audi��o
                        tn['Surdo-cegueira'] # audi��o e vis�o
                        ), axis=1)

    deficiencia = ['Cegueira', 'Baixa Vis�o', 
                    'Surdez', 'Auditiva', 
                    'F�sica', 
                    'Surdo-cegueira', 
                    'M�ltipla', 
                    'Mental', 'Autismo infantil', 'S�ndrome de Asperger', 'S�ndrome de Rett', 'Transtorno Desintegrativo da Inf�ncia', 
                    'Altas habilidades/superdota��o']
    for nome_deficiencia in deficiencia:
        df_alunos[nome_deficiencia] = df_alunos.apply(lambda d: 
            valida_campo_deficiencia(d['Tem Necessidade'], d[nome_deficiencia]), axis=1)

    # Remove os espa�os do in�cio e final das c�lulas, caso haja algum
    df_alunos[df_alunos.columns] = df_alunos.apply(lambda x: x.str.strip())

    # Cria o arquivo dados e escreve a primeira linha com 40|10
    dados = open("dados_alunos.txt", "w+", encoding='utf-8')
    # TODO: Ver se isso se manter� assim
    dados.write("40|10\n")

    # Escreve as informa��es no xt
    escreve_txt(df_alunos, dados)

if __name__ == '__main__':
    # Chamada da fun��o main
    main()

# Cabe�alho do arquivo - Registro 40: A princ�pio OK
# Registro do aluno - Registro 41: N�o OK
# Ok > 2, 4 , 5, 12 - 21
# Meio OK > 7 (Separar a string do ano)
# Problema de com 2 colunas: 8 (Precisa de outro doc), 9 (Precisa de outro doc), 10 , 11
# ?: 11
# D�vidas > (3) Que base? Dever�amos comparar com outra?
# Registro do v�nculo do aluno com o curso - Registro 42: N�o OK
# Ok > 6

# Corrigir as linhas que informam o Total de Registro e outras que n�o informam dados de alunos