import openpyxl
import ProcesingData as PS
import tkinter as tk

def ExtractInfo():
   wb = openpyxl.load_workbook('E:/Descargas/julio 2024.xlsx')
   ws = wb.active
   columDataNames = txtname.get()
   columDataPreach = txtpreach.get()
   columDataCourses = txtcourses.get()
   columDataNotes = txtnotes.get()
   starDataNames = txtstar.get()
   endDataNames = txtend.get()
   month = txtmonth.get()

   if listPublicador.curselection() == "Publicador":
       while starDataNames <= endDataNames:
         name = ws[columDataNames + str(starDataNames)].value
         preach = ws[columDataPreach + str(starDataNames)].value
         courses = ws[columDataCourses + str(starDataNames)].value
         notes = ws[columDataNotes + str(starDataNames)].value + ' '
         name = str(name.upper())
         PS.PDF_RowSetNormal(name, month, preach, courses, notes)
         starDataNames = starDataNames + 1
   elif listPublicador.curselection() == "Precursor auxiliar":
      while starDataNames <= endDataNames:
         name = ws[columDataNames + str(starDataNames)].value
         courses = ws[columDataCourses + str(starDataNames)].value
         notes = ws[columDataNotes + str(starDataNames)].value + ' '
         hours = ws[columDataPreach + str(starDataNames)].value
         name = str(name.upper())
         PS.PDF_RowSetAux(name, month, courses, hours, notes)
         starDataNames = starDataNames + 1
   elif listPublicador.curselection() == "Precursor regular":
      while starDataNames <= endDataNames:
         name = ws[columDataNames + str(starDataNames)].value
         courses = ws[columDataCourses + str(starDataNames)].value
         notes = ws[columDataNotes + str(starDataNames)].value + ' '
         hours = ws[columDataPreach + str(starDataNames)].value
         name = str(name.upper())
         PS.PDF_RowSetColporter(name, month, courses, hours, notes)
         starDataNames = starDataNames + 1

ventana = tk.Tk()
ventana.geometry("800x600+100+50")
ventana.title("Auto PDF tarjetas predicacion")

lblname = tk.Label(ventana, text="Columna Nombre")
lblname.pack()
txtname = tk.Entry(ventana, width=30)
txtname.pack()
lblpreach = tk.Label(ventana, text="Columna Predicacion")
lblpreach.pack()
txtpreach = tk.Entry(ventana, width=30)
txtpreach.pack()
lblcourses = tk.Label(ventana, text="Columna Cursos")
lblcourses.pack()
txtcourses = tk.Entry(ventana, width=30)
txtcourses.pack()
lblnotes = tk.Label(ventana, text="Columna Notas")
lblnotes.pack()
txtnotes = tk.Entry(ventana, width=30)
txtnotes.pack()
lblstar = tk.Label(ventana, text="Celda inicial")
lblstar.pack()
txtstar = tk.Entry(ventana, width=30)
txtstar.pack()
lblend = tk.Label(ventana, text="Celda final")
lblend.pack()
txtend = tk.Entry(ventana, width=30)
txtend.pack()
lblmonth = tk.Label(ventana, text="Mes")
lblmonth.pack()
txtmonth = tk.Entry(ventana, width=30)
txtmonth.pack()

listPublicador = tk.Listbox(ventana, width=30, selectmode="single")
listPublicador.pack()
publicadores = ["Publicador", "Precursor auxiliar", "Precursor regular"]
for publicador in publicadores:
    listPublicador.insert(tk.END, publicador)

btniniciar = tk.Button(ventana, text="Iniciar grabado", command=ExtractInfo)
btniniciar.pack()

ventana.mainloop()

