'''
Script recebe um arquivo Excel  da trilha de certificação dos cursos e também recebe um Excel dos alunos e as disciplinas que ele faz parte, 
retornando no final o nome do Aluno e os certificados intermediários que ele recebeu

Desenvolvido por Vinicius Tozo
Última atualização: 17/07/2020
'''

# coding: utf-8
import pandas

arquivo_regras = "../in/Base estudantes Trilhas Integradas por curso.xlsx"
df_regras = pandas.read_excel(arquivo_regras)

arquivo_disciplinas = "../in/relatorio_arquivo_exportado.xlsx"
df_alunos_disciplinas = pandas.read_excel(arquivo_disciplinas)

dict_aluno = dict()
dict_certificado = dict()

for index, line in df_alunos_disciplinas.iterrows():
    key = str(line["Codigo Completo"]) + ";" + line["Nome Pessoa"] + ";" + \
          str(line["Curso Aluno"]) + ";" + line["Nome Curso Aluno"]
    if key in dict_aluno:
        dict_aluno[key].append(line["Nome Disciplina Oferta"].lower())
    else:
        dict_aluno[key] = [line["Nome Disciplina Oferta"].lower()]

for index, line in df_regras.iterrows():
    key = str(line["Curso"]) + ";" + line["Certificação"]
    if key in dict_certificado:
        dict_certificado[key].append(line["Disciplina"].lower())
    else:
        dict_certificado[key] = [str(line["Disciplina"]).lower()]

for aluno in dict_aluno:
    curso_aluno = aluno.split(";")[2]
    for certificado in dict_certificado:
        curso_certificado = certificado.split(";")[0]
        if curso_aluno != curso_certificado:
            continue
        apto = True
        for disciplina in dict_certificado[certificado]:
            if disciplina not in dict_aluno[aluno]:
                apto = False
                # print("Erro", aluno, certificado, disciplina)
        if apto:
            print(str(aluno) + ";" + certificado)
