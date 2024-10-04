from tkinter import Tk
from tkinter.filedialog import askopenfilename
import pandas as pd
import string as s


def main(arquivo):
    qtd = []
    discente = []

    qtAluno = 0

    for nome in range(len(arquivo['NOME_ALUNO'])):
        nome = s.capwords(arquivo['NOME_ALUNO'][nome])
        if nome not in discente:
            discente.append(nome)
            qtd.append(1)
            qtAluno += 1
        else:
            index = discente.index(nome)
            qtd[index] += 1

    for n in range(len(discente)):
        print(f'Discente: {discente[n]} > Qtd de Registros: {qtd[n]}')
    print(f'Qtd de alunos: {qtAluno}')

    # Data e Situação do discente
    d_sit_d = []

    last_name = ''
    for line in range(len(arquivo['NOME_ALUNO'])):
        alun_name = s.capwords(arquivo['NOME_ALUNO'][line])
        if last_name != alun_name:

            print(line, alun_name)
            last_name = alun_name
    # Oq demoraria mais?
    # 22131 registros passando por 45638 dados cada ou
    # 45638 dados verificando 22131 itens cada


if __name__ == '__main__':
    print('Selecione o relatório para tratamento dos vinculos')

    Tk().withdraw()
    vinculos_alunos = askopenfilename(
        filetypes=[('Arquivo excel', '.xlsx')],
        title='Selecione relatório de vinculos')
    print(f'{vinculos_alunos}')

    arquivo = pd.read_excel(vinculos_alunos, sheet_name='Exportar Planilha', dtype=str)

    main(arquivo)

