from tkinter import Tk
from tkinter.filedialog import askopenfilename
import pandas as pd
import string as s
import time as t
from datetime import datetime


# TODO: Fazer o tratamento das trocas de grade sem marcação
def calcula_tempo_trancamento(dtt):
    periodos = 0  # Seta para 0 a qtd de períodos trancados
    hoje = datetime.today().strftime('%Y-%m')  # Pega a data atual
    mes_atual, ano_atual = int(hoje.split('-')[-1]), int(hoje.split('-')[0])  # Separa o ano e o mês atual

    curso = dtt['CURSO_ID'][0]

    for d in range(len(dtt['DT_SIT_ACADEMICA'])):
        dataSit = dtt['DT_SIT_ACADEMICA'][d].split(' ')[0]  # Pega a data do registro da stuação acadêmica
        mes, ano = int(dataSit.split('-')[1]), int(dataSit.split('-')[0])  # Pega ano e mês do registro

        try:  # Tenta verificar se o próximo valor existe
            # Se existe verifica se o valor atual indica o trancamento e o próximo indica reabertura
            # Para encerrar a contagem
            if curso != dtt['CURSO_ID'][d]:
                periodos = 0
                curso = dtt['CURSO_ID'][d]

            if ((dtt['SIT_ACADEMICA'][d] == 'Trancado' or
                dtt['SIT_ACADEMICA'][d] == 'Trancamento Institucional') and
                    dtt['SIT_ACADEMICA'][d + 1] == 'Entrada via Reabertura'):
                dataSit_re = dtt['DT_SIT_ACADEMICA'][d + 1].split(' ')[0]
                mes_re, ano_re = int(dataSit_re.split('-')[1]), int(dataSit_re.split('-')[0])
                periodos += int((((ano_re - ano) * 12) + (mes_re - mes)) / 6) + 1
        except:
            # Se não conseguir fazer a verificação ou o último valor estiver com trancamento
            # Calcula a quantidade de períodos até o semestre atual
            if (dtt['SIT_ACADEMICA'][d] == 'Trancado' or
                    dtt['SIT_ACADEMICA'][d] == 'Trancamento Institucional'):
                periodos += int((((ano_atual - ano) * 12) + (mes_atual - mes)) / 6) + 1

    return periodos


# Gera o arquivo Excel com Nome Discente, Número Períodos com trancamento,
# Última Situação Acadêmica e Data da última situação
def gerar_excel(n_d, n_p, n_s, n_ds):

    n_df = pd.DataFrame(data={'Discente': n_d,
                              'Última Situação': n_s,
                              'Data Útlima Situação': n_ds,
                              'Períodos com Trancamento': n_p})
    n_df.to_excel('relatório_trancamentos.xlsx', index=False)


def main(arquivo):
    discente = []  # Nomes dos alunos para geração do mini dataframe
    qtd = []  # Ocorrências do mesmo aluno

    n_discente = []  # Nomes dos alunos para inserir na planilha
    n_periodos = []  # Qtd de períodos com trancamento
    n_ultima_sit = []  # Última situação acadêmica registrada
    n_data_ultima_st = []  # Data que a última situação acadêmica foi registrada

    for nome in range(len(arquivo['NOME_ALUNO'])):
        # nome = s.capwords(arquivo['NOME_ALUNO'][nome])
        nome = arquivo['NOME_ALUNO'][nome]
        if nome not in discente:
            discente.append(nome)
            qtd.append(1)
        else:
            index = discente.index(nome)
            qtd[index] += 1

    start = t.time()
    alunos_distintos = 0
    for n in range(len(discente)):
        '''if n != 1079:
            continue'''
        # Gera um novo dataframe temporário com apenas os dados do aluno
        dtt = arquivo[arquivo['NOME_ALUNO'].str.contains(discente[n])].sort_values(by='DT_SIT_ACADEMICA').reset_index()
        # dtt = dtt.sort_values(by='DT_SIT_ACADEMICA').reset_index()
        print(dtt)

        # Tenta calcular a qtd de tempo que o aluno ficou com o tranacamento
        try:
            tempo_t = calcula_tempo_trancamento(dtt)
        except:
            tempo_t = 'Indeterminado'
        print(f'{s.capwords(discente[n])} trancado a {tempo_t} períodos\n'
              # f'Última situação acadêmica: {dtt["ULTIMA_SIT_ACAD"][dtt.index[-1]]}'
              )

        n_discente.append(s.capwords(discente[n]))
        n_periodos.append(tempo_t)
        try:
            n_ultima_sit.append(dtt['ULTIMA_SIT_ACAD'][dtt.index[-1]])
        except:
            n_ultima_sit.append('Sem Registro')
        try:
            n_data_ultima_st.append(dtt['DT_ULTIMA_SIT_ACAD'][dtt.index[-1]])
        except:
            n_data_ultima_st.append('Sem Registro')

        alunos_distintos += 1
        '''if alunos_distintos == 25:
            break'''
    print(alunos_distintos)
    end = t.time()
    print(f'Tempo de leitura: {end - start}')
    # chama a função de geração do Excel
    gerar_excel(n_discente, n_periodos, n_ultima_sit, n_data_ultima_st)


if __name__ == '__main__':
    print('Selecione o relatório para tratamento dos vinculos')

    Tk().withdraw()
    vinculos_alunos = askopenfilename(
        filetypes=[('Arquivo excel', '.xlsx')],
        title='Selecione relatório de vinculos')
    print(f'{vinculos_alunos}')

    arquivo = pd.read_excel(vinculos_alunos, sheet_name='Exportar Planilha', dtype=str)

    main(arquivo)
