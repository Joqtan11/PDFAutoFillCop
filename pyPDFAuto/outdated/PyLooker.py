from PyPDF2 import PdfReader

file = 'E:/Publicadores//1 ADERITA DE ARCE.pdf'
pdfReader = PdfReader(file)
Checkboxes = pdfReader.get_fields()
print(Checkboxes)
