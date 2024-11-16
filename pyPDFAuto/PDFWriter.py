import os
import pdfrw
from tkinter import messagebox
ANNOT_KEY = '/Annots'           # key for all annotations within a page
ANNOT_FIELD_KEY = '/T'          # Name of field. i.e. given ID of field
ANNOT_FORM_type = '/FT'         # Form type (e.g. text/button)
ANNOT_FORM_button = '/Btn'      # ID for buttons, i.e. a checkbox
ANNOT_FORM_text = '/Tx'         # ID for textbox
SUBTYPE_KEY = '/Subtype'
WIDGET_SUBTYPE_KEY = '/Widget'

def write_fillable_pdf(input_pdf_path, output_pdf_path, data_dict):
    
    
    try:
        template_pdf = pdfrw.PdfReader(input_pdf_path)

        for Page in template_pdf.pages:
            if Page[ANNOT_KEY]:
                for annotation in Page[ANNOT_KEY]:
                    if annotation[ANNOT_FIELD_KEY] and annotation[SUBTYPE_KEY] == WIDGET_SUBTYPE_KEY :
                        key = annotation[ANNOT_FIELD_KEY][1:-1] # Remove parentheses
                        if key in data_dict.keys():
                            if annotation[ANNOT_FORM_type] == ANNOT_FORM_button:
                                # button field i.e. a checkbox
                                annotation.update(pdfrw.PdfDict(V=data_dict[key], AS=pdfrw.PdfName(data_dict[key])))
                            elif annotation[ANNOT_FORM_type] == ANNOT_FORM_text:
                                # regular text field
                                annotation.update( pdfrw.PdfDict(V=data_dict[key], AS=data_dict[key]))
        template_pdf.Root.AcroForm.update(pdfrw.PdfDict(NeedAppearances=pdfrw.PdfObject('true')))
        pdfrw.PdfWriter().write(output_pdf_path, template_pdf)
    except: FileNotFoundError(print(messagebox.showinfo(message='no se pudo encontrar el archivo: ' + input_pdf_path, title='Error'))) 
