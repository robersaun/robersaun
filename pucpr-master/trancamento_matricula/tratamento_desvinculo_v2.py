from tkinter import Tk
from tkinter.filedialog import askopenfilename
import pandas as pd
import string as s
import time as t


def calcula_tempo_trancamento(dtt):
    periodos = 0

    for d in range(len(dtt['DT_SIT_ACADEMICA'])):
        if len(dtt['DT_SIT_ACADEMICA']) == 1:
            # TODO: considerar apenas o dado único, se tiver mais q 1, considerar o próximo para validar se há reabertura
            print('Tenho apenas 1 dado')
        pass

    return periodos


def main(arquivo):
    dados_discente = []

    qtd = []
    discente = []

    for nome in range(len(arquivo['NOME_ALUNO'])):
        # nome = s.capwords(arquivo['NOME_ALUNO'][nome])
        nome = arquivo['NOME_ALUNO'][nome]
        if nome not in discente:
            discente.append(nome)
            qtd.append(1)
        else:
            index = discente.index(nome)
            qtd[index] += 1

    """    
    for line in range(len(arquivo['NOME_ALUNO'])):
        nome = s.capwords(arquivo['NOME_ALUNO'][line])
        dados_discente.append([nome, arquivo['SIT_ACADEMICA'][line],
                               str(arquivo['DT_SIT_ACADEMICA'][line]).split(' ')[0],
                               arquivo['ULTIMA_SIT_ACAD'][line],
                               str(arquivo['DT_ULTIMA_SIT_ACAD'][line]).split(' ')[0]])
    """

    start = t.time()
    alunos_distintos = 0
    for n in range(len(discente)):
        # Gera um novo dataframe temporário com apenas os dados do aluno
        dtt = arquivo[arquivo['NOME_ALUNO'].str.contains(discente[n])]
        dtt = dtt.sort_values(by='DT_SIT_ACADEMICA')
        print(dtt)

        # Calcula a qtd de tempo que o aluno ficou com o tranacamento
        tempo_t = calcula_tempo_trancamento(dtt)

        alunos_distintos += 1
        if alunos_distintos == 25:
            break
    print(alunos_distintos)
    end = t.time()
    print(f'Tempo de leitura: {end - start}')


if __name__ == '__main__':
    print('Selecione o relatório para tratamento dos vinculos')

    Tk().withdraw()
    vinculos_alunos = askopenfilename(
        filetypes=[('Arquivo excel', '.xlsx')],
        title='Selecione relatório de vinculos')
    print(f'{vinculos_alunos}')

    arquivo = pd.read_excel(vinculos_alunos, sheet_name='Exportar Planilha', dtype=str)

    main(arquivo)
