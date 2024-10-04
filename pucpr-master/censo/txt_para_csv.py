'''
Converte um aquivo txt dos dados pessoais do aluno para um arquivo csv (tabela)

Desenvolvido por Vinicius Tozo
Última atualização: 03/01/2021
'''


# -*- coding: utf-8 -*-
from tkinter import Tk
from tkinter.filedialog import askopenfilename

Tk().withdraw()
filename = askopenfilename()
txt = open(filename, encoding="utf-8")
dados_pessoais = open("dados_pessoais.csv", "w")
dados_vinculos = open("dados_vinculos.csv", "w")

dados_vinculos.write(
    'ID do Aluno;Nome do Aluno;Tipo de registro;Semestre de referência;Código do Curso;Código do pólo do '
    'curso a distância ;ID na IES - Identificação única do aluno na IES;Turno do aluno;Situação de '
    'vínculo do aluno ao curso;Curso origem;Semestre de conclusão do curso;Aluno PARFOR;Semestre de '
    'ingresso no curso;Tipo de escola que concluiu o Ensino Médio;"Forma de ingresso/seleção  - '
    'Vestibular";"Forma de ingresso/seleção  - Enem";"Forma de ingresso/seleção  - Avaliação '
    'Seriada";"Forma de ingresso/seleção  - Seleção Simplificada";"Forma de ingresso/seleção  - Egresso '
    'BI/LI";"Forma de ingresso/seleção  - PEC-G";"Forma de ingresso/seleção  - Transferência Ex '
    'Officio";"Forma de ingresso/seleção  - Decisão judicial";"Forma de ingresso  - Seleção para Vagas '
    'Remanescentes";"Forma de ingresso  - Seleção para Vagas de Programas Especiais";Mobilidade '
    'acadêmica;Tipo de mobilidade acadêmica;IES destino;Tipo de mobilidade acadêmica internacional;País '
    'destino;Programa de reserva de vagas;"Programa de reserva de vagas/açoes afirmativas - '
    'Etnico";"Programa de reserva de vagas/ações afirmativas - Pessoa com deficiência";"Programa de '
    'reserva de vagas - Estudante procedente de escola pública";"Programa de reserva de vagas/ações '
    'afirmativas - Social/renda familiar";"Programa de reserva de vagas/ações afirmativas - '
    'Outros";Financiamento estudantil;"Financiamento Estudantil Reembolsável - FIES";"Financiamento '
    'Estudantil Reembolsável - Governo Estadual";"Financiamento Estudantil Reembolsável -Governo '
    'Municipal";"Financiamento Estudantil Reembolsável  - IES";"Financiamento Estudantil Reembolsável - '
    'Entidades externas";"Tipo de financiamento não reembolsável - ProUni integral";"Tipo de '
    'financiamento não reembolsável - ProUni parcial";"Tipo de financiamento não reembolsável - Entidades '
    'externas";"Tipo de financiamento não reembolsável - Governo estadual";"Tipo de financiamento não '
    'reembolsável - IES";"Tipo de financiamento não reembolsável - Governo municipal";Apoio Social;"Tipo '
    'de apoio social - Alimentação";"Tipo de apoio social - Moradia";"Tipo de apoio social - '
    'Transporte";"Tipo de apoio social - Material didático";"Tipo de apoio social - Bolsa trabalho";"Tipo '
    'de apoio social - Bolsa permanência";Atividade extracurricular;"Atividade extracurricular - '
    'Pesquisa";"Bolsa/remuneração referente à atividade extracurricular - Pesquisa";"Atividade '
    'extracurricular - Extensão";"Bolsa/remuneração referente à atividade extracurricular - '
    'Extensão";"Atividade extracurricular - Monitoria";"Bolsa/remuneração referente à atividade '
    'extracurricular - Monitoria";"Atividade extracurricular - Estágio não '
    'obrigatório";"Bolsa/remuneração referente à atividade extracurricular - Estágio não '
    'obrigatório";Carga horária total do curso por aluno;Carga horária integralizada pelo aluno\n')

id_aluno = ""
nome_aluno = ""

for line in txt:
    if line.startswith("31") or line.startswith("41"):  # Linha de docente ou aluno
        values = line.split("|")
        id_aluno = str(values[1])
        nome_aluno = values[2]

        csv_separator = ";"
        csv_values = csv_separator.join(values)
        dados_pessoais.write(csv_values)
    elif line.startswith("32") or line.startswith("42"):
        values = line.split("|")

        csv_separator = ";"
        csv_values = csv_separator.join([id_aluno] + [nome_aluno] + values)
        dados_vinculos.write(csv_values)

dados_pessoais.close()
dados_vinculos.close()
