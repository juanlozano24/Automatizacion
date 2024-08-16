import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
#import workbook

ventana = tk.Tk()
ventana.geometry("700x700")
ventana.title("Modulo automatización")
ventana.pack_propagate(False)
ventana.resizable(0, 0)
ventana.iconbitmap(True, "C:/Users/JuanL/automat/source/icons/icons.ico")

frame_ini = ttk.Label(ventana, text="Esta Herramienta permite realizar los cruces de informacion \nnecesarios para obtener las tecnologias que seran objeto de automatización", justify=['center'])
frame_ini.place(relx=0.21, rely=0.01)

#Frame o cuadro de dialogo inicial donde se carga el archivo de anexo evento
frame_1 = tk.LabelFrame(ventana, text="1. Cargar Anexo Evento", labelanchor="n")
frame_1.place(height=80, width=650, rely=0.08, relx=0, x = 25)

btn_cargue_1 = tk.Button(frame_1, text="Visualizar archivo",command=lambda: cargue_anexo_evento())
btn_cargue_1.place(rely=0.1,relx=0.55)

btn_selec_1 = tk.Button(frame_1, text="Seleccionar archivo",command=lambda: nombre_archivo())
btn_selec_1.place(rely=0.1,relx=0)

label_file_1 = ttk.Label(frame_1, text="No se ha seleccionado un archivo.xlsx")
label_file_1.place(rely=0.15, relx=0.18)

#Frame o cuadro de dialogo donde se carga la maestra de tecnologias suceptibles de automatizacion actualizada
frame_2 = tk.LabelFrame(ventana, text="2. Cargar maestra de Tecnologias suceptibles de automatizacíon",labelanchor='n')
frame_2.place(height=80, width=650, rely=0.21, relx=0, x=25)

btn_cargue_2 = tk.Button(frame_2, text="Visualizar archivo",command=lambda: cargue_maestra())
btn_cargue_2.place(rely=0.1,relx=0.15)

btn_selec_2 = tk.Button(frame_2, text="Seleccionar archivo",command=lambda: nombre_archivo())
btn_selec_2.place(rely=0.1, relx=0)

label_file_2 = ttk.Label(frame_2, text="No se ha seleccionado un archivo.xlsx")
label_file_2.place(rely=0.15, relx=0.18)

#Frame o cuadro donde se dupura el archivo anexo evento
frame_3 = tk.LabelFrame(ventana, text="3. Depuración del archivo anexo", labelanchor='n')
frame_3.place(height=100, width=200, rely=0.35, relx=0, x=25)

btn_depurar = tk.Button(frame_3, text="Depurar archivo anexo", command=lambda: depuracion_anexo())
btn_depurar.place(rely=0.55, relx=0.18)

#Frame o cuadro donde se ejecuta cruce de anexo evento vs maestra tecnologias suceptibles para automatización
frame_4 = tk.LabelFrame(ventana, text="4. AnexoEventoVsMaestraTecnologias", labelanchor='n')
frame_4.place(height=100, width=250, rely=0.35, relx=0.3, x=25)

btn_cruce_1 = tk.Button(frame_4, text="Realizar cruce", command=lambda: evento_vs_maestra())
btn_cruce_1.place(rely=0.55, relx=0.35)

#Frame o cuadro de vista del resultado del archivo cargado
frame_vista = tk.LabelFrame(ventana, text="Vista previa del archivo cargado",labelanchor='n')
frame_vista.place(height=200, width=700, rely=0.65)

tv1 = ttk.Treeview(frame_vista)
tv1.place(relheight=1, relwidth=1) 

treescrolly = tk.Scrollbar(frame_vista, orient="vertical", command=tv1.yview) # command means update the yaxis view of the widget
treescrollx = tk.Scrollbar(frame_vista, orient="horizontal", command=tv1.xview) # command means update the xaxis view of the widget
tv1.configure(xscrollcommand=treescrollx.set, yscrollcommand=treescrolly.set) # assign the scrollbars to the Treeview Widget
treescrollx.pack(side="bottom", fill="x") # make the scrollbar fill the x axis of the Treeview widget
treescrolly.pack(side="right", fill="y") # make the scrollbar fill the y axis of the Treeview widget

def nombre_archivo():
    filename = filedialog.askopenfilename(initialdir="/",
                                          title="Seleccionar un Archivo.xlsx",
                                          filetype=(("xlsx files", "*.xlsx"),("All Files", "*.*")))
    label_file_1["text"] = filename

def cargue_anexo_evento():
    input_cols = [0,1,5,6,9,13]
    file_path = label_file_1["text"]
    try:
        excel_filename = r"{}".format(file_path)
        if excel_filename[-4:] == ".csv":
            df = pd.read_csv(excel_filename, usecols=input_cols)
        else:
            df = pd.read_excel(excel_filename, usecols=input_cols)

    except ValueError:
        tk.messagebox.showerror("Advertencia", "El archivo cargado es invalido")
        return None
    except FileNotFoundError:
        tk.messagebox.showerror("Error", f"No se ha cargardo ningun archivo {file_path}")
        return None

    clear_data()
    tv1["column"] = list(df.columns)
    tv1["show"] = "headings"
    for column in tv1["columns"]:
        tv1.heading(column, text=column)

    df_rows = df.to_numpy().tolist()
    for row in df_rows:
        tv1.insert("", "end", values=row)
    return df

def cargue_maestra():
    """filename = filedialog.askopenfilename(initialdir="/",
                                          title="Seleccionar un Archivo.xlsx",
                                          filetype=(("xlsx files", "*.xlsx"),("All Files", "*.*")))
    label_file_2["text"] = filename"""
    input_cols1 = [0]
    file_path = label_file_1["text"]
    try:
        excel_filename = r"{}".format(file_path)
        if excel_filename[-4:] == ".csv":
            df_m = pd.read_csv(excel_filename, sheet_name="NO SE HACE", usecols=input_cols1)
        else:
            df_m = pd.read_excel(excel_filename, sheet_name="NO SE HACE", usecols=input_cols1)

    except ValueError:
        tk.messagebox.showerror("Advertencia", "El archivo cargado es invalido")
        return None
    except FileNotFoundError:
        tk.messagebox.showerror("Error", f"No se ha cargardo ningun archivo {file_path}")
        return None
    return df_m

def depuracion_anexo():
    dfanx = cargue_anexo_evento()
    dfanx = dfanx[dfanx["TARIFA NEGOCIADA*"] <= 3000000]
    dfanx.sort_values(by=["TARIFA NEGOCIADA*"], ascending = False)
    dfanx = dfanx[dfanx["ESTADO"] == "Activo"]
    dfanx = dfanx.drop_duplicates(['COD TECNOLOGIA* (RIPS)', 'DESC TECNOLOGIA*'], keep='last')

    def convertir_a_numero(value):
        try:
            return int(value)
        except ValueError:
            return value

    dfanx['COD TECNOLOGIA* (RIPS)'] = dfanx['COD TECNOLOGIA* (RIPS)'].apply(convertir_a_numero)

    clear_data()
    tv1["column"] = list(dfanx.columns)
    tv1["show"] = "headings"
    for column in tv1["columns"]:
        tv1.heading(column, text=column) # let the column heading = column name

    df_rows = dfanx.to_numpy().tolist() # turns the dataframe into a list of lists
    for row in df_rows:
        tv1.insert("", "end", values=row) # inserts each list into the treeview. For parameters see https://docs.python.org/3/library/tkinter.ttk.html#tkinter.ttk.Treeview.insert
    return dfanx

def evento_vs_maestra():
    
    dfm = cargue_maestra()
    dfa = depuracion_anexo()
    
    criterio_no = dfm.merge(dfa, how='left', left_on='COD TECNOLOGIA* (RIPS)', right_on='cod_tecnologia', indicator='Criterio_no')
    
    clear_data()
    tv1["column"] = list(criterio_no.columns)
    tv1["show"] = "headings"
    for column in tv1["columns"]:
        tv1.heading(column, text=column) # let the column heading = column name

    df_rows = criterio_no.to_numpy().tolist() # turns the dataframe into a list of lists
    for row in df_rows:
        tv1.insert("", "end", values=row) # inserts each list into the treeview. For parameters see https://docs.python.org/3/library/tkinter.ttk.html#tkinter.ttk.Treeview.insert
    return None

def clear_data():
    tv1.delete(*tv1.get_children())
    return None

ventana.mainloop()
