'''
Script realiza a formatação de arquivo do tipo xls para um arquivo do tipo xlsx

Desenvolvido por Vinicius Tozo
Última atualização: 13/08/2021
'''

import win32com.client as win32

fname = "C:\\Users\\Vini\\Documents\\GitHub\\estagio-pucpr\\acao-matrizes\\Grade_curricular_CBJD 2021_1.xls"
excel = win32.gencache.EnsureDispatch('Excel.Application')
wb = excel.Workbooks.Open(fname)

wb.SaveAs(fname + "x", FileFormat=51)   # FileFormat = 51 is for .xlsx extension
wb.Close()                              # FileFormat = 56 is for .xls extension
excel.Application.Quit()
