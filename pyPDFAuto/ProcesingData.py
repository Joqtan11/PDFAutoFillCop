import PDFWriter as PDF
import AsignMonth as As
def PDF_RowSetNormal(Pdf_name, month, PreachCheck, courses, notes):
    PDFNameImput = 'E:/Publicadores/' + Pdf_name + '.pdf'
    PDFNameoutput = 'E:/SF/' + Pdf_name + '.pdf'
    if PreachCheck == 'No':
        PreachCheck = 'Off'
    else:
        PreachCheck = 'Yes' 
    month = int(As.AsignMonth(month))

    if notes != "":
            DataFile = {
                '901_'+ str(month) +'_CheckBox': str(PreachCheck),
                '902_'+ str(month) +'_Text_C_SanSerif': str(courses),
                '905_'+ str(month) +'_Text_SanSerif': notes
            }
    else:
            DataFile = {
                '901_'+ str(month) +'_CheckBox': str(PreachCheck),
                '902_'+ str(month) +'_Text_C_SanSerif': str(courses),
            }
    PDF.write_fillable_pdf(PDFNameImput, PDFNameoutput, DataFile)

def PDF_RowSetColporter(Pdf_name, month, courses, hours, notes):
    PDFNameImput = 'E:/Publicadores/PRECURSORES REGULARES/' + Pdf_name + '.pdf'
    PDFNameoutput = 'E:/SF/PRECURSORES REGULARES/' + Pdf_name + '.pdf'

    month = int(As.AsignMonth(month))

    if notes != "":
            DataFile = {
                '901_'+ str(month) +'_CheckBox': 'Yes',
                '902_'+ str(month) +'_Text_C_SanSerif': str(courses),
                '904_'+ str(month) +'_S21_Value': str(hours),
                '905_'+ str(month) +'_Text_SanSerif': notes
            }
    else:
            DataFile = {
                '901_'+ str(month) +'_CheckBox': 'Yes',
                '902_'+ str(month) +'_Text_C_SanSerif': str(courses),
                '904_'+ str(month) +'_S21_Value': str(hours),
            }
    PDF.write_fillable_pdf(PDFNameImput, PDFNameoutput, DataFile)

def PDF_RowSetAux(Pdf_name, month, courses, hours, notes):
    PDFNameImput = 'E:/Publicadores/PRECURSORES AUXILIARES/' + Pdf_name + '.pdf'
    PDFNameoutput = 'E:/SF/PRECURSORES AUXILIARES/' + Pdf_name + '.pdf'

    month = int(As.AsignMonth(month))

    if notes != "":
        DataFile = {
                '901_'+ str(month) +'_CheckBox': 'Yes',
                '902_'+ str(month) +'_Text_C_SanSerif': str(courses),
                '903_'+ str(month) +'_CheckBox': 'Yes',
                '904_'+ str(month) +'_S21_Value': str(hours) ,
                '905_'+ str(month) +'_Text_SanSerif': notes
      }
    else:
        DataFile = {
                '901_'+ str(month) +'_CheckBox': 'Yes',
                '902_'+ str(month) +'_Text_C_SanSerif': str(courses),
                '903_'+ str(month) +'_CheckBox': 'Yes',
                '904_'+ str(month) +'_S21_Value': str(hours) ,
         }
    PDF.write_fillable_pdf(PDFNameImput, PDFNameoutput, DataFile)
