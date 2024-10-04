import os
import pandas
import re
from tkinter import filedialog
from tkinter import Tk

def unifica(root):
    key = 0
    dict_unificado={}
    # Para cada diretório dentro do diretório raiz
    for dirpath, dirnames, filenames in os.walk(root):
        # Se o diretório possuir outros diretórios dentro dele, ignora
        if dirnames:
            continue
        for file in filenames:
            caminho = f"{dirpath}/{file}"
            df = pandas.read_excel(caminho)
            df.fillna('',inplace=True)
            for dados in df.iterrows():
                df.columns.values[16] = 'Qual Ambiente de Aprendizagem?(Laboratório)'
                if dados[1]['Escola'] == '':
                    break
                else:
                    dict_unificado[key] = {
                        'Escola': dados[1]["Escola"],
                        'Cód. do Curso':dados[1]["Cód. do Curso"],
                        'Nome do Curso': dados[1]["Nome do Curso"],
                        'Cód. da Disciplina': dados[1]["Cód. da Disciplina"],
                        'Nome da disciplina': dados[1]["Nome da disciplina"],
                        'Período': dados[1]["Período"],
                        'Turma': dados[1]["Turma"],
                        'Matriz': dados[1]["Matriz"],
                        'Turno': dados[1]["Turno"],
                        'Presencial': dados[1]["Presencial"],
                        'Tipo de Atividade': dados[1]["Tipo de Atividade"],
                        'Modulação': dados[1]["Modulação"],
                        'Modulação pela Matriz': dados[1]["Modulação pela Matriz"],
                        'Carga Horária': dados[1]["Carga Horária"],
                        'Previsão do número de alunos': dados[1]["Previsão do número de alunos"],
                        'Usa Sala Teórica, Metodologia Ativa ou Laboratório?': dados[1]["Usa Sala Teórica, Metodologia Ativa ou Laboratório?"],
                        'Qual Ambiente de Aprendizagem?(Laboratório)': dados[1]['Qual Ambiente de Aprendizagem?(Laboratório)'],
                        'Software (Especificar nomes dos softwares)': dados[1]["Software (Especificar nomes dos softwares)"],
                        'Possui unificação com outro curso? (Indicar o curso/turma)': dados[1]["Possui unificação com outro curso? (Indicar o curso/turma)"],
                        'Indicar curso base/pagante em caso de unificação.': dados[1]["Indicar curso base/pagante em caso de unificação."],
                        'Docente 2022.2':dados[1]["Docente 2022.2"],
                        'Indicar como eletiva?': dados[1]["Indicar como eletiva?"],
                        'Curso/Escola/Instituição?': dados[1]["Curso/Escola/Instituição?"],
                        'Observação': dados[1]["Observação"],

                    }
                    key+=1
    new_df = pandas.DataFrame(data=dict_unificado)

    new_df = (new_df.T)
    new_df.to_excel('Arquivo_Unificado.xlsx', sheet_name='Template', index=False)





if __name__ == '__main__':
    Tk().withdraw()
    pasta_raiz = filedialog.askdirectory()
    unifica(pasta_raiz)
