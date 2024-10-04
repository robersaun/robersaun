"""Simulador de Carga Horária Docente

Script para gerar uma planilha contendo o simulador de carga horária docente.
Recebe como entrada o arquivo de Carga Horária Consolidada em formato Excel,
e gera um simulador de carga horária docente também em formato Excel.

Solicitado pela equipe de informações acadêmicas, não foi utilizado.
"""
import pandas
from tkinter import Tk
from tkinter.filedialog import askopenfilename


def formata_worksheet(workbook, worksheet, skiprows, ultima_linha, sheet_name):
    # Cria formatações para as células
    unlocked = workbook.add_format({'locked': False})
    hidden = workbook.add_format({'hidden': True})
    user_input = workbook.add_format({'bg_color': '#E09358'})

    # Para cada linha, escreve as fórmulas
    for row in range(skiprows + 2, ultima_linha + 2):
        # Fórmula de Total
        formula = f'=F{row}+G{row}'
        worksheet.write_formula(f"H{row}", formula, hidden)

        # Fórmula de Regime de Trabalho
        formula = f'=_xlfn.IFS(' \
                  f'H{row}>40,"Excedeu a Carga Horária",' \
                  f'H{row}=0,"Sem Carga Horária",' \
                  f'AND(H{row}=40,G{row}/H{row}>=0.5),"Tempo Integral",' \
                  f'AND(H{row}>=36,G{row}/H{row}>=0.5,C{row}="Curitiba"),"Tempo Integral",' \
                  f'AND(H{row}>=36,G{row}/H{row}>=0.5,C{row}="Toledo"),"Tempo Integral",' \
                  f'AND(H{row}>=12,G{row}/H{row}>=0.25),"Tempo Parcial",' \
                  f'TRUE, "Horista"' \
                  f')'
        worksheet.write_formula(f"I{row}", formula, hidden)

    # Habilita a proteção da planilha, menos colunas E e F
    worksheet.protect(options={'autofilter': True})
    worksheet.set_column('F:G', None, unlocked)
    worksheet.data_validation(f'F{skiprows + 2}:G{ultima_linha + 2}', {'validate': 'decimal',
                                                                       'criteria': 'between',
                                                                       'minimum': 0,
                                                                       'maximum': 40,
                                                                       'input_title': 'Digite a carga horária:',
                                                                       'input_message': 'entre 0 e 40'})

    # Formata a largura das colunas
    worksheet.set_column('A:A', 12)
    worksheet.set_column('B:B', 48)
    worksheet.set_column('C:C', 24)
    worksheet.set_column('D:D', 24)
    worksheet.set_column('E:E', 24)
    worksheet.set_column('F:F', 19)
    worksheet.set_column('G:G', 19)
    worksheet.set_column('H:H', 19)
    worksheet.set_column('I:I', 24)

    # Formata como tabela
    worksheet.add_table(skiprows, 0, ultima_linha, 8, {'style': 'Table Style Medium 17', 'columns': [
        {'header': 'Código RH'},
        {'header': 'Professor'},
        {'header': 'Campus'},
        {'header': 'Escola'},
        {'header': 'Curso'},
        {'header': 'Horas Letivas'},
        {'header': 'Horas Não Letivas'},
        {'header': 'Total'},
        {'header': 'Regime de Trabalho'}
    ]})

    # Aplica formatação diferente no cabeçalho para as colunas editáveis
    worksheet.write(f'F{skiprows + 1}', 'Horas Letivas', user_input)
    worksheet.write(f'G{skiprows + 1}', 'Horas Não Letivas', user_input)

    # Inclui a tabela de totais
    regimes_trabalho = ["Tempo Integral", "Tempo Parcial", "Horista", "Excedeu a Carga Horária",
                        "Sem Carga Horária"]
    for i in range(0, len(regimes_trabalho)):
        worksheet.write(f"C{i + 4}", regimes_trabalho[i])
        formula = f'=SUMPRODUCT(' \
                  f'SUBTOTAL(' \
                  f'3,OFFSET(I{skiprows + 2}:I{ultima_linha + 1},' \
                  f'ROW(I{skiprows + 2}:I{ultima_linha + 1})-ROW(I{skiprows + 2}),0,1)),' \
                  f'--(I{skiprows + 2}:I{ultima_linha + 1}=C{i + 4}))'
        worksheet.write_formula(f"D{i + 4}", formula, hidden)

    worksheet.write(f'C9', 'Total')
    worksheet.write_formula(f'D9', '=SUM(D4:D8)', hidden)

    worksheet.add_table(2, 2, 8, 3, {'style': 'Table Style Medium 19', 'columns': [
        {'header': 'Regime de Trabalho'},
        {'header': 'Total'}
    ]})

    # Adiciona um gráfico
    chart = workbook.add_chart({'type': 'pie'})
    chart.add_series({
        'values': '=\'' + sheet_name + '\'!D4:D8',
        'categories': '=\'' + sheet_name + '\'!C4:C8',
        'data_labels': {'value': True, 'percentage': True, 'leader_lines': True, 'font': {'size': 16}}
    })
    worksheet.insert_chart('G1', chart, {'x_scale': 0.9, 'y_scale': 0.9})

    # Adiciona uma imagem
    worksheet.insert_image('A1', 'logo.jpg')
    worksheet.set_row(0, 66)

    # Adiciona o título
    header = workbook.add_format({
        'bold': 1,
        'align': 'center',
        'valign': 'vcenter',
        'font_color': '#8A0538',
        'font_size': 24
    })
    worksheet.merge_range('A1:F1', 'SIMULADOR DE REGIME DE TRABALHO', header)

    # Remove as linhas de grade
    worksheet.hide_gridlines(2)


def main(arquivo):
    # Lê o arquivo e seleciona as colunas
    dataframe = pandas.read_excel(arquivo, sheet_name="BD_CH_CONSOLIDADA", engine="openpyxl")
    dataframe = dataframe[
        ["CÓDIGO RH", "PROFESSOR", "CAMPUS VÍNCULO", "ESCOLA VÍNCULO", "CURSO VÍNCULO", "TIPO", "QTDE", "ENTIDADE"]]

    # Identifica as horas letivas
    dataframe['TIPO'] = dataframe['TIPO'].replace([
        'Letiva',
        'Tutoria Online'
    ], 'Letivas')

    # Identifica as horas não letivas
    dataframe['TIPO'] = dataframe['TIPO'].replace([
        'Não letiva',
        'Tutoria Online',
        'Orientação Matriz Curricular',
        'Permanência de Disciplina Especial',
        'Erro de Cadastro'
    ], 'Não Letivas')

    # Renomeia as escolas de vínculo
    dataframe['ESCOLA VÍNCULO'] = dataframe['ESCOLA VÍNCULO'].str.replace('PUCPR-ESCOLA DE ', "")
    dataframe['ESCOLA VÍNCULO'] = dataframe['ESCOLA VÍNCULO'].str.replace('PUCPR-ESCOLA ', "")
    dataframe['ESCOLA VÍNCULO'] = dataframe['ESCOLA VÍNCULO'].str.replace('PUCPR-', "")
    dataframe['ESCOLA VÍNCULO'] = dataframe['ESCOLA VÍNCULO'].str.capitalize()
    dataframe['CAMPUS VÍNCULO'] = dataframe['CAMPUS VÍNCULO'].str.capitalize()
    # Renomeia as escolas da atividade
    dataframe['ENTIDADE'] = dataframe['ENTIDADE'].str.upper()
    dataframe['ENTIDADE'] = dataframe['ENTIDADE'].str.replace('PUCPR-ESCOLA DE ', "")
    dataframe['ENTIDADE'] = dataframe['ENTIDADE'].str.replace('PUCPR-ESCOLA ', "")
    dataframe['ENTIDADE'] = dataframe['ENTIDADE'].str.replace('PUCPR-', "")
    dataframe['ENTIDADE'] = dataframe['ENTIDADE'].str.capitalize()

    # Corrige a escola dos campus fora de sede
    dataframe.loc[dataframe['CAMPUS VÍNCULO'].eq('Londrina'), 'ESCOLA VÍNCULO'] = 'Londrina'
    dataframe.loc[dataframe['CAMPUS VÍNCULO'].eq('Maringa'), 'ESCOLA VÍNCULO'] = 'Maringá'
    dataframe.loc[dataframe['CAMPUS VÍNCULO'].eq('Maringá'), 'ESCOLA VÍNCULO'] = 'Maringá'
    dataframe.loc[dataframe['CAMPUS VÍNCULO'].eq('Toledo'), 'ESCOLA VÍNCULO'] = 'Toledo'

    # Cria dataframe com a ligação entre professor e entidades
    dataframe_professor_entidade = dataframe[["CÓDIGO RH", "ENTIDADE"]]
    dataframe_professor_entidade = dataframe_professor_entidade.drop_duplicates()

    # Transforma os valores da coluna 'TIPO' em colunas 'Horas Letivas' e
    # 'Horas Não Letivas', calculando a soma de horas para cada professor
    dataframe = dataframe.groupby([
        "CÓDIGO RH", "PROFESSOR", "CAMPUS VÍNCULO", "ESCOLA VÍNCULO", "CURSO VÍNCULO", "TIPO"
    ]).sum().squeeze().unstack().add_prefix('Horas ').reset_index()
    dataframe = dataframe.sort_values(by=['PROFESSOR'], ascending=True)
    dataframe = dataframe.drop_duplicates("CÓDIGO RH")

    # Preenche os valores vazios com zero
    dataframe["Horas Letivas"] = dataframe["Horas Letivas"].fillna(0)
    dataframe["Horas Não Letivas"] = dataframe["Horas Não Letivas"].fillna(0)

    # Seleciona os diferentes valores para Escola
    lista_escolas = dataframe["ESCOLA VÍNCULO"]
    lista_escolas = lista_escolas.drop_duplicates()
    lista_escolas = lista_escolas.values.tolist()

    skiprows = 11

    # Para cada escola
    for escola in lista_escolas:
        # Cria o dataframe filtrado e o arquivo de saída
        dataframe_por_vinculo = dataframe.loc[dataframe["ESCOLA VÍNCULO"] == escola]
        nome_arquivo_saida = f"Simulador {escola}.xlsx"
        arquivo_saida = pandas.ExcelWriter(path=nome_arquivo_saida, engine="xlsxwriter")
        sheet_name = "Professores vinculados"
        dataframe_por_vinculo.to_excel(arquivo_saida, sheet_name=sheet_name, index=False, startrow=skiprows)

        # Cria e seleciona uma planilha
        workbook = arquivo_saida.book
        worksheet = arquivo_saida.sheets[sheet_name]
        ultima_linha = len(dataframe_por_vinculo.index) + skiprows

        formata_worksheet(workbook, worksheet, skiprows, ultima_linha, sheet_name)

        # Segunda aba
        # Cria o dataframe filtrado e o arquivo de saída
        dataframe_professor_entidade_escola = dataframe_professor_entidade.loc[
            dataframe_professor_entidade["ENTIDADE"] == escola]
        dataframe_por_atividade = dataframe.loc[
            dataframe["CÓDIGO RH"].isin(dataframe_professor_entidade_escola["CÓDIGO RH"])]
        # Se nenhum professor foi encontrado passa pra próxima escola
        if len(dataframe_por_atividade.index) == 0:
            arquivo_saida.save()
            continue
        sheet_name = "Professores com CH na escola"
        dataframe_por_atividade.to_excel(arquivo_saida, sheet_name=sheet_name, index=False, startrow=skiprows)

        # Cria e seleciona uma planilha
        workbook = arquivo_saida.book
        worksheet = arquivo_saida.sheets[sheet_name]
        ultima_linha = len(dataframe_por_atividade.index) + skiprows

        formata_worksheet(workbook, worksheet, skiprows, ultima_linha, sheet_name)

        # Salva o arquivo
        arquivo_saida.save()


if __name__ == '__main__':
    Tk().withdraw()
    nome_arquivo = askopenfilename(
        filetypes=[('Arquivo excel', '.xlsx')],
        title='Selecione o arquivo de Carga Horária Consolidada')

    main(nome_arquivo)
