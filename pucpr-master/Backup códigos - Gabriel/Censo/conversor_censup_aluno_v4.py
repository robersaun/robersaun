# -*- coding: Windows-1252 -*-

from tkinter import Tk
from tkinter.filedialog import askopenfilename
import pandas as pd

# Pega o ano que est� considerando para fazer a divis�o
# Compara para ver se � menor
# ant_post verifica se o valor para comparar deve ser anterior ou posterior
def verificar_ano_ingresso(ano, ingresso, ant_post):
    valores = []
    pega_ano = ''
    for n in ingresso:
        valores.append(n)
    for n in range(len(valores)):
        if n != 0:
            pega_ano += valores[n]
    ano = int(ano)
    pega_ano = int(pega_ano)
    if ant_post == 'posterior':
        if pega_ano > ano:
            return True
        else:
            return False
    # Aparentemente essa parte do c�digo n�o ser� usada pois ambas, 
    # at� o momento, aparecem como posterior
    elif ant_post == 'anterior':
        if pega_ano < ano:
            return True
        else:
            return False

# Cod.: A04 -- Fun��o para verificar a validade e a exist�ncia do CPF
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

# Cod.: A05 -- Verifica se o aluno � estrangeiro ou n�o
def verifica_estrangeiro(docE, nacionalidade):
    nacionalidade = str(nacionalidade)
    docE = str(docE)
    if nacionalidade == '':
        return "Nacionalidade n�o informada"
    elif nacionalidade == '1' or nacionalidade == '2':
        return "Preenchimento exclusivo para alunos estrangeiros"
    else:
        return docE

# Cod.: A07 -- verifica cor e ra�a do estudante
# TODO: 2� etapa da valida��o
# TODO: Ver onde mostrar os avisos
def verifica_cr(cr, ingresso):
    info = ''
    ano = '2014'
    ant_post = 'posterior'
    posterior = False
    tamData = 5
    if cr == "":
        info = "Cor/Ra�a n�o informada"
    elif cr in "0123456":
        if cr == "0":
            info = 'Aluno n�o quis declarar cor/ra�a'
        elif cr == "1":
            info = 'Branca'
        elif cr == "2":
            info = 'Preta'
        elif cr == "3":
            info = 'Parda'
        elif cr == "4":
            info = 'Amarela'
        elif cr == "5":
            info = 'Ind�gena'
        elif cr == "6":
            info = 'N�o disp�e da informa��o'
    else:
        info = "Preencher com um dos valore: 0, 1, 2, 3, 4, 5, 6"
    # >> Ver como validar nessa situa��o
    # >> Aparentemente em alguns casos s� tem o semestre ou est� vazio
    if tamData == len(ingresso):
        posterior = verificar_ano_ingresso(ano, ingresso, ant_post)
    if (posterior == True and cr == '6'):
        return 'A op��o "N�o disp�es da informa��o" n�o pode ser informada para ingressantes com ano de ingresso posteriores a 2014'
    else:
        return info

# Cod.: A12 -- Verifica se o aluno possui alguma defici�ncia
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

# Cod.: A1321 -- Verifica se o campo da defici�ncia foi preenchido
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

# Cod.: A22 -- Verifica qual o foi o tipo de Ensino M�dio que o estudante se formou
# TODO: Verificar sobre o Curso Origem
# TODO: Verificar oq fazer quando o ano n�o for identificado
def verifica_em(em, ingresso, nacionalidade):
    tipo = ''
    ano = '2012' # Para validar o ano igual a 2013
    ant_post = 'posterior'
    posterior = False
    tamData = 5
    if em == "0":
        tipo = "Privado"
    elif em == "1":
        tipo = "Publico"
    elif em == "2":
        tipo = "N�o disp�e da informa��o"
    else:
        tipo = "Preencher com um dos valores: 0, 1, 2"
    # >> Ver oq deve ocorrer se a nacionalidade n�o for informada
    # >> Ver sobre a data do curso de origem
    if (em=='2' and (nacionalidade=='1' or nacionalidade=='2')):
        if tamData == len(ingresso):
            posterior = verificar_ano_ingresso(ano, ingresso, ant_post)
        if posterior==True:
            return "A op��o 'N�o disp�e da informa��o' somente poder� ser informada para v�nculo com ano de ingresso menor que 2013 ou para alunos estrangeiros"
        else:
            return tipo
    else:
        return tipo

# Cod.: C06 -- Verifica qual foi o turno selecionado
# TODO: 2� etapa de valida��o
# TODO: Ver como verificar qual a modalidade do curso
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

# Cod.: C07 -- Verifica Situa��o
# TODO: Ver oq fazer quanto a data de conclus�o >  para verificar a idade do aluno
# TODO: Ver como q pega o tipo do curso
def verifica_situacao(situacao):
    situacao_vinculo = ''
    if situacao in "234567":
        if situacao == "2":
            situacao_vinculo = "Cursando"
        if situacao == "3":
            situacao_vinculo = "Matr�cula trancada"
        if situacao == "4":
            situacao_vinculo = "Desvinculado do curso"
        if situacao == "5":
            situacao_vinculo = "Transferido para outro curso da mesma IES"
        if situacao == "6":
            situacao_vinculo = "Formado"
        if situacao == "7":
            situacao_vinculo = "Falecido"
        return situacao_vinculo
    else:
        return "Preencher com um dos valores: 2,3,4,5,6,7"

# Cod.: C10 -- Verifica PARFOR
#TODO: pegar grau acad�mico
def verifica_parfor(parfor, vaga_especial):
    tipo_parfor = ''
    # >> Verificar oq deve ocorrer se PARFOR for vazio 
    if parfor == "":
        tipo_parfor = "N�o informado"
    if parfor == "0":
        tipo_parfor = "N�o"
    elif parfor == "1":
        tipo_parfor = "Sim"
    else:
        tipo_parfor = "Preencher com um dos valores: (vazio),0,1"
    if (parfor == "1" and vaga_especial == "0"):
        return "Aluno PARFOR exige vaga de ingresso sele��o para vagas de programas especiais"
    else:
        return tipo_parfor

# Cod.: GED -- Para cada discente, escreve os dados pessoais
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

    '''--------------------------------------------------------------------------------
        Leitura das p�ginas do arquivo
        -------------------------------------------------------------------------------'''
    # Planilha de alunos � armazenado na vari�vel conforme a escola 
    # Sendo respectivamente: escola de direito, belas artes, educa��o e humanidade ou medicina
    df_direito = pd.read_excel(planilha_alunos, dtype=str, sheet_name="Direito")
    df_belas = pd.read_excel(planilha_alunos, dtype=str, sheet_name="Belas")
    df_eh = pd.read_excel(planilha_alunos, dtype=str, sheet_name="EH")
    df_medicina = pd.read_excel(planilha_alunos, dtype=str, sheet_name="Medicina")

    # A princ�pio tem que tratar as linhas que s� tem as informa��es de totais
    # Talvez influencie na quest�o do curso > Para pegar o nome do curso
    # Concatena as p�ginas, em uma �nica
    df_alunos = pd.concat([df_direito, df_belas, df_eh, df_medicina])

    # Remove os nan das c�lulas vazias
    df_alunos = df_alunos.fillna("")

    '''--------------------------------------------------------------------------------
        Leitura das informa��es contidas nas colunas do arquivo
        --------------------------------------------------------------------------------'''

    # ----- Informa��es dos alunos ------

    # Cod.: A04 -- Chama a fun��o de tratar CPF para cada CPF de cada tabela de alunos
    df_alunos['CPF'] = df_alunos['CPF'].apply(lambda cpf: trata_cpf(cpf))

    # Cod.: A05 -- Verifica se h� documento de estrangeiro
    df_alunos['Documento Estrangeiro'] = df_alunos.apply(lambda n:
        verifica_estrangeiro(n['Documento Estrangeiro'], n['Nacionalidade']), axis=1)

    # Cod.: A07 -- Verifica cor e ra�a
    df_alunos['Ra�a'] = df_alunos.apply(lambda cr: 
        verifica_cr(cr['Ra�a'], cr['Semestre de Ingresso']), axis=1)

    # Cod.: A12 -- Verificar alunos com defici�ncia
    df_alunos['Tem Necessidade'] = df_alunos.apply(lambda tn: 
        tem_necessidade(tn['Tem Necessidade'], # gen�rico
                        tn['Cegueira'], tn['Baixa Vis�o'], # relacionados a vis�o
                        tn['Surdez'], tn['Auditiva'], # relacionado a audi��o
                        tn['Surdo-cegueira'] # audi��o e vis�o
                        ), axis=1)

    # Cod.: A1321 -- Verifica qual/quais as defici�ncias do aluno
    deficiencia = ['Cegueira', 'Baixa Vis�o', 
                    'Surdez', 'Auditiva', 
                    'F�sica', 
                    'Surdo-cegueira', 
                    'M�ltipla', 
                    'Mental', 
                    'Autismo infantil', 'S�ndrome de Asperger', 'S�ndrome de Rett', 'Transtorno Desintegrativo da Inf�ncia', 
                    'Altas habilidades/superdota��o'] # falta intelectual(?), tgd/tda
    for nome_deficiencia in deficiencia:
        df_alunos[nome_deficiencia] = df_alunos.apply(lambda d: 
            valida_campo_deficiencia(d['Tem Necessidade'], d[nome_deficiencia]), axis=1)

    # Cod.: A22 -- Verifica local de conclus�o do Ensino M�dio
    df_alunos['Tipo Escola'] = df_alunos.apply(lambda em: 
        verifica_em(em['Tipo Escola'], em['Semestre de Ingresso'], em['Nacionalidade']), axis=1)

    # ----- Informa��es para v�nculo dos alunos com o curso ------

    # Cod.: C06 -- Verifica Turno
    df_alunos['Turno'] = df_alunos['Turno'].apply(lambda turno: verifica_turno(turno))

    # Cod.: C07 -- Verifica Situa��o
    df_alunos['Situa��o'] = df_alunos['Situa��o'].apply(lambda situacao: verifica_situacao(situacao))

    # Cod.: C10 -- Verifica PARFOR
    df_alunos['PARFOR'] = df_alunos.apply(lambda parfor: 
        verifica_parfor(parfor['PARFOR'], parfor['Vagas de Programas Especiais']), axis=1)

    # Cod.: C14
    forma_ingresso = ['Vstibular' , 'Enem', 'Avalia��o Seriada', 'Sele��o Simplificada', 'Egresso BI/LI', ' PEC-G', 'Transfer�ncia Ex Officio', 'Decis�o Judicial']

    ''' --------------------------------------------------------------------------------
        Escrita das informa��es no .txt
        --------------------------------------------------------------------------------'''

    # Remove os espa�os do in�cio e final das c�lulas, caso haja algum
    df_alunos[df_alunos.columns] = df_alunos.apply(lambda x: x.str.strip())

    # Cria o arquivo dados e escreve a primeira linha com 40|10
    dados = open("dados_alunos.txt", "w+", encoding='utf-8')
    # TODO: Ver se isso se manter� assim
    dados.write("40|10\n")

    # Cod.: GED -- Escreve as informa��es no xt
    escreve_txt(df_alunos, dados)

if __name__ == '__main__':
    # Chamada da fun��o main
    main()

# Cabe�alho do arquivo - Registro 40: A princ�pio OK

# Registro do aluno - Registro 41: N�o OK
# Ok > 2, 4 , 5, 12 - 21
# Meio OK >
# ?: 11, 8 (Precisa de outro doc), 9 (Precisa de outro doc), 10
# D�vidas > 3 e 6 (Que base? Dever�amos comparar com outra?), 20 (� gen�rico para todos da mesma categoria?)

# 6 (N�o tem o formato da data - falta fazer as valida��es conforme os anos) >> Data sempre ddmmaaaa

# Registro do v�nculo do aluno com o curso - Registro 42: N�o OK
# Ok > 
# Meio ok > 6, 7, 10 (forma de ingresso/sele��o?), 7 ("verificar idade do aluno")
# ? > 11 - 13 (no 13 qual deve ser o formato da data?)

# Pq tem algumas linhas com a escrita em vermelho?
# Corrigir as linhas que informam o Total de Registro e outras vazias
