from tkinter import ttk, Tk, Label, Button, filedialog
from tkinter import *

def fileDialog(lbl_txtbx):
    filename = filedialog.askopenfilename(initialdir = "/", title = "Seleccionar archivo ...", filetype = (("xls", "*.xls"), ("xlsx", "*xlsx"), ("All Files", "*.*")))
    lbl_txtbx.configure(text = filename)

window = Tk()
window.title('Complemento ARMS: CSV Generator')
window.geometry("650x380")

# Generate instructions 1st step: locate XLS file
lbl_instructions_locate_file = Label(window, text="(1) Seleccionar el archivo '*.xls' con la programación de ARMS:")
lbl_instructions_locate_file.grid(column = 0, row = 1, padx = 20, pady = 20)

# Generate TextBox
lbl_txtbx = Label(window, text="Seleccionar Archivo ...", bg="white")
lbl_txtbx.grid(column = 0, row = 2, padx = 20, pady = 20)

# Generate Browse Button
btn_Browse = Button(window, text='Buscar Archivo ...', command = fileDialog(lbl_txtbx))
btn_Browse.grid(column = 0, row = 3, padx = 20, pady = 20)

# Generate instructions 2st step: generate .csv file
lbl_instructions_generate_csv = Label(window, text="(2) Una vez ubicado el arhivo, generar el '*.csv':")
lbl_instructions_generate_csv.grid(column = 0, row = 4, padx = 20, pady = 20)

# Generate 'Crear CSV' Button


window.mainloop()
