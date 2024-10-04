import pandas as pd
import xlrd
import datetime
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import xlsxwriter
import pytz

dicionario = {
    "Acompanhamento e Orientação Acadêmica" : "Manutenção de CH cfe Cláusula 29a. da CCT",
    "Administrativas" : "Administrativas",
    "Avaliação de Atividades Complementares" : "Manutenção de CH cfe Cláusula 29a. da CCT",
    "CCH - Coordenadoria de Carga Horária" : "Manutenção de CH cfe Cláusula 29a. da CCT",
    "CEP - Comitê de Ética em Pesquisa" : "Manutenção de CH cfe Cláusula 29a. da CCT",
    "CEUA - Comitê de Ética no Uso de Animais" : "Manutenção de CH cfe Cláusula 29a. da CCT",
    "Construção de Portfólio por Competência" : "Manutenção de CH cfe Cláusula 29a. da CCT",
    "Coordenação - Complementação" : "Coordenação - Portaria",
    "Coordenação - Portaria" : "Coordenação - Portaria",
    "Coordenação Adjunta - Portaria" : "Coordenação - Portaria",
    "Coordenação Adjunta Unidade Acadêmica - Portaria" : "Coordenação - Portaria",
    "CrEAre - Centro de Ensino e Aprendizagem" : "Manutenção de CH cfe Cláusula 29a. da CCT",
    "Direção - Portaria" : "Coordenação - Portaria",
    "EAD" : "Manutenção de CH cfe Cláusula 29a. da CCT",
    "Eixos Temáticos de Disciplinas" : "Manutenção de CH cfe Cláusula 29a. da CCT",
    "Estágio curricular" : "Manutenção de CH cfe Cláusula 29a. da CCT",
    "HNB - Programa Habilidades do Núcleo Básico" : "Manutenção de CH cfe Cláusula 29a. da CCT",
    "Licença" : "Licença",
    "Manutenção de CH cfe Cláusula 29a. da CCT" : "Manutenção de CH cfe Cláusula 29a. da CCT",
    "Monitoria" : "Manutenção de CH cfe Cláusula 29a. da CCT",
    "NDE - Núcleo Docente Estruturante" : "Manutenção de CH cfe Cláusula 29a. da CCT",
    "NDH - Núcleo de Direitos Humanos" : "Manutenção de CH cfe Cláusula 29a. da CCT",
    "NEO - Núcleo Empregabilidade e Oportunidades" : "Manutenção de CH cfe Cláusula 29a. da CCT",
    "NEP - Núcleo de Excelência Pedagógica" : "Manutenção de CH cfe Cláusula 29a. da CCT",
    "NPJ - Núcleo de Práticas Jurídicas" : "Manutenção de CH cfe Cláusula 29a. da CCT",
    "NPP - Núcleo de Práticas em Psicologia" : "Manutenção de CH cfe Cláusula 29a. da CCT",
    "Núcleo de Internacionalização" : "Manutenção de CH cfe Cláusula 29a. da CCT",
    "Organização de Eventos Acadêmicos" : "Manutenção de CH cfe Cláusula 29a. da CCT",
    "Organização do ENADE" : "Manutenção de CH cfe Cláusula 29a. da CCT",
    "PIBEP" : "Manutenção de CH cfe Cláusula 29a. da CCT",
    "PIBIC" : "Manutenção de CH cfe Cláusula 29a. da CCT",
    "PIBIC Jr" : "Manutenção de CH cfe Cláusula 29a. da CCT",
    "PIBIC MASTER" : "Manutenção de CH cfe Cláusula 29a. da CCT",
    "PIBIC MOBILIDADE NACIONAL" : "Manutenção de CH cfe Cláusula 29a. da CCT",
    "PIBITI" : "Manutenção de CH cfe Cláusula 29a. da CCT",
    "Planejamento de ações de engajamento e organização das atividades dos estudantes visando a melhoria dos "
    "indicadores de aprendizagem e retenção" : "Manutenção de CH cfe Cláusula 29a. da CCT",
    "Portfólio de Educação Continuada" : "Manutenção de CH cfe Cláusula 29a. da CCT",
    "Preceptoria em Programa de Aprimoramento Profissional em Medicina Veterinária" : "Manutenção de CH cfe Cláusula 29a. da CCT",
    "Prescrição Médica" : "Manutenção de CH cfe Cláusula 29a. da CCT",
    "Produção de Material Didático" : "Manutenção de CH cfe Cláusula 29a. da CCT",
    "Projeto de Gestão Acadêmica" : "Manutenção de CH cfe Cláusula 29a. da CCT",
    "Projeto de Pesquisa e Desenvolvimento - Especificar projeto" : "Manutenção de CH cfe Cláusula 29a. da CCT",
    "Projeto Interdisciplinar" : "Manutenção de CH cfe Cláusula 29a. da CCT",
    "Psicopedagógica" : "Manutenção de CH cfe Cláusula 29a. da CCT",
    "Responsável Estágio Curricular" : "Manutenção de CH cfe Cláusula 29a. da CCT",
    "Responsável pelo Estágio não Obrigatório" : "Manutenção de CH cfe Cláusula 29a. da CCT",
    "Responsável por TCC" : "Manutenção de CH cfe Cláusula 29a. da CCT",
    "Simulação Clinica" : "Manutenção de CH cfe Cláusula 29a. da CCT",
    "Stricto Sensu" : "Stricto Sensu",
    "TCC/Monografia" : "Manutenção de CH cfe Cláusula 29a. da CCT",
    "Tutorial (Medicina/Saude)" : "Manutenção de CH cfe Cláusula 29a. da CCT",
    "Unidades da Educação e Humanidades" : "Manutenção de CH cfe Cláusula 29a. da CCT",
    "Unidades da Medicina" : "Manutenção de CH cfe Cláusula 29a. da CCT",
    "Unidades de Arquitetura e Design" : "Manutenção de CH cfe Cláusula 29a. da CCT",
    "Unidades de Ciências da Vida" : "Manutenção de CH cfe Cláusula 29a. da CCT",
    "Unidades de Comunicação e Artes" : "Manutenção de CH cfe Cláusula 29a. da CCT",
    "Unidades de Engenharia e Computação" : "Manutenção de CH cfe Cláusula 29a. da CCT",
    "Unidades de Londrina" : "Manutenção de CH cfe Cláusula 29a. da CCT",
    "Unidades de Maringá" : "Manutenção de CH cfe Cláusula 29a. da CCT",
    "Unidades de Negócios" : "Manutenção de CH cfe Cláusula 29a. da CCT",
    "Unidades de Toledo" : "Manutenção de CH cfe Cláusula 29a. da CCT",
    "Unidades do Direito" : "Manutenção de CH cfe Cláusula 29a. da CCT"
}

Tk().withdraw()
relacao_alunos = askopenfilename(
    filetypes=[('Arquivo excel', '.xlsx')],
    title='Selecione a relação de alunos')

df_alunos = pd.read_excel(relacao_alunos, index=False)

del df_alunos['Matricula']
del df_alunos['Campus']
del df_alunos['Escola']
del df_alunos['Mês']
del df_alunos['Ano']
del df_alunos['CR Campus']
del df_alunos['CR Escola']
del df_alunos['Centro Resultado']
del df_alunos['Natureza']
del df_alunos['Grupo Atividade']
del df_alunos['Grupo de Análise']
del df_alunos['Classe de Análise']
del df_alunos['Unidade Acadêmica']
del df_alunos['Disciplina/Turma']

df_alunos.insert(4, 'CR Efetivo', df_alunos['CR'].tolist())

df_alunos['CR'].tolist()

df_alunos.rename(columns={'CR': 'CR Fonte'}, inplace=True)

df_alunos.insert(5, 'Atividade2', df_alunos['Atividade'].tolist())
df_alunos.insert(5, 'Atividade2', df_alunos['Atividade'].tolist())

df_alunos['Atividade'] = df_alunos['Atividade'].map(dicionario)

df_alunos = df_alunos[df_alunos['Tipo Atividade'].eq('Não Letiva')]

df_alunos = df_alunos[df_alunos['Atividade2'].ne('Orientação MC - Matriz Curricular')]
df_alunos = df_alunos[df_alunos['Atividade2'].ne('DE - Projeto Integrador')]
df_alunos = df_alunos[df_alunos['Atividade2'].ne('Construção de Portfólio por Competência ')]

del df_alunos['Tipo Atividade']

df_alunos.rename(columns={'Atividade2': 'Atividade'}, inplace=True)

utc_now = pytz.utc.localize(datetime.datetime.utcnow())
dt_now = utc_now.astimezone(pytz.timezone("America/Sao_Paulo"))

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('arquivo_importacao ' + str(dt_now.strftime("%Y-%m-%d %H-%M-%S")) + '.xlsx', engine='xlsxwriter')

# Convert the dataframe to an XlsxWriter Excel object.
df_alunos.to_excel(writer, sheet_name='Sheet1', index=False)

# Close the Pandas Excel writer and output the Excel file.
writer.save()


