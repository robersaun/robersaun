import PyPDF2

FILE_PATH = 'relatorio_bookings.pdf'

with open(FILE_PATH, mode='rb') as f:
    reader = PyPDF2.PdfFileReader(f)
    page = reader.getPage(0)
    print(page.extractText())
