from numpy import append
import pandas as pd
import re
from tkinter import Tk
from tkinter.filedialog import askopenfilename


# Busca no texto a palavra
def verificar_protocolos(protocolo, texto):
    cont = []
    for i in range(len(protocolo)):
        cont.append(False)
        try:
            if re.search('\\b'+protocolo[i]+'\\b', texto, re.IGNORECASE):
                cont[i] = True
        except:
            continue
    return cont


# Permite a seleção do arquivo
Tk().withdraw()
nome_arquivo = askopenfilename(
    title="Selecione o arquivo",
    filetypes=[("Excel files", ".xlsx")])

# Lê o arquivo
dataframe = pd.read_excel(nome_arquivo, sheet_name="Filtrado",
    names=['Nº Protocolo', 'Comentários'], header=0, dtype={'Nº Protocolo': str, 'Comentários': str})

# Tipo de protocolo
# Se precisar adicionar mais é inserir aqui
# Para colocar as palavras com escrita diferente e o mesmo significado,
# colocar tudo na mesma String separado por | sem espaço
protocolo = ["Equivalência",
             "Leitura e Escrita Acadêmica|LEA",
             "Matricula|Matricular|Matrícula",
             "Incluir|Inclusão",
             "Excluir|Exclusão",
             "Eletiva",
             "Boleto",
             "Mensalidade",
             "Valor errado|Incorreto",
             "FIES",
             "PROUNI",
             "Bolsa",
             "Dispensa|Dispensada|Dispensado",
             "Vaga|Vagas",
             "Conflito de horário",
             "Sem oferta",
             "Disciplina não aparece|Sem disciplina",
             "Turma",
             "Turno"
             ]

# Faz a contagem dos protocolos
contagem_protocolo = []
for i in range(len(protocolo)):
    contagem_protocolo.append(0)

linha = 0
# Busca cada um dos textos para realizar a leitura
for i in dataframe['Comentários']:
    texto = dataframe['Comentários'][linha]
    contem_protocolo = verificar_protocolos(protocolo, texto)
    linha += 1
    #Incrementa a quantidade
    for j in range(len(protocolo)):
        if contem_protocolo[j]:
            contagem_protocolo[j] += 1

# Cria uma lista para organizar a prioridade do maior para o menor
priorizacao_protocolo = []
for i in range(len(protocolo)):
    priorizacao_protocolo.append([0, 0])

# Insere o protocolo com o valor para organizar em ordem de quantidade em uma lista interna
for i in range(len(protocolo)):
    priorizacao_protocolo[i][0] = protocolo[i]
    priorizacao_protocolo[i][1] = contagem_protocolo[i]

# Faz o reajuste da lista conforme a quantidade, do maior para o menor
priorizacao_protocolo.sort(key=lambda q: q[1], reverse=True)

# Imprime os resultados
print("------------------------")
for i in range(len(protocolo)):
    nome_protocolo = priorizacao_protocolo[i][0]
    quantidade_protocolo = priorizacao_protocolo[i][1]
    # Realiza um corte no texto, para mostrar apenas uma das palavras junto com o total
    nome_protocolo = nome_protocolo.split("|")
    nome_protocolo = nome_protocolo[0]
    print(f"{nome_protocolo}: {quantidade_protocolo}")
print("------------------------")
