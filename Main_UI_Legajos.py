
import os
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, messagebox
import shutil
from PIL import Image
import comtypes.client
from pdf_utils import unir_pdfs
from zip_utils import crear_zip_legajos

#------------CONSTANTES -------------
EXTENSIONES_WORD = (".doc", ".docx")
EXTENSIONES_EXCEL = (".xls", ".xlsx")
EXTENSIONES_IMG = (".png", ".jpg", ".jpeg")
EXTENSION_PDF = ".pdf"

#-------------FUNCIONES -----------

def seleccionar_origen():
    carpeta = filedialog.askdirectory(title="Seleccionar carpeta origen")
    if carpeta:
        ruta_origen.set(carpeta)

def seleccionar_destino():
    carpeta = filedialog.askdirectory(title="Seleccionar carpeta destino")
    if carpeta:
        ruta_destino.set(carpeta)

def generar_legajos():
    origen = ruta_origen.get()
    destino = ruta_destino.get()

    if not origen or not destino:
        messagebox.showerror(
            "Error",
            "Debes seleccionar la carpeta origen y la carpeta destino"
        )
        return
    
    carpeta_colaboradores = []

    # Caso 1: carpeta origen contiene subcarpetas
    subcarpetas = [
        os.path.join(origen, nombre)
        for nombre in os.listdir(origen)
        if os.path.isdir(os.path.join(origen, nombre))
    ]

    if subcarpetas:
        carpeta_colaboradores = subcarpetas
    else:
        #Caso 2: es una sola carpeta
        carpeta_colaboradores.append(origen)

    resumen = []

    #-------PARA BARRA DE PROGRESO---
    progress["value"] = 0
    progress["maximum"] = len(carpeta_colaboradores)
    root.update_idletasks()
 
    #---------------------------------------------

    for carpeta in carpeta_colaboradores:
        nombre_carpeta = os.path.basename(carpeta)
        dni, nombre = extraer_dni_nombre(nombre_carpeta)

        if dni is None:
            resumen.append(f"{nombre_carpeta} formato inválido")
            continue

        pdfs, convertibles, ignorados = clasificar_archivos(carpeta)

        carpeta_temp = crear_carpeta_temporal(carpeta)

        # copiar PDFs originales
        copiar_archivos(carpeta, pdfs, carpeta_temp)
        
        # convertir NO PDFs
        convertidos = 0
        for archivo in convertibles:
            ruta_archivo = os.path.join(carpeta, archivo)
            convertir_a_pdf_guardar(ruta_archivo, carpeta_temp)
            convertidos += 1

        resumen.append(
            f"{dni} - {nombre}\n"
            f" PDFs Originales copiados: {len(pdfs)}\n"
            f" Archivos convertidos a PDF: {len(convertibles)}\n"
            f" Carpeta temp: solo PDFs"
        )

        #Unir PDFs
        pdf_final = os.path.join(
            ruta_destino.get(), f"{dni} - {nombre}.pdf"
        )
        unir_pdfs(carpeta_temp, pdf_final)

        #--------PARA BARRA DE PROGRESO---------
        progress["value"] += 1
        root.update_idletasks()
        #------------------------------

    #crear ZIP
    ruta_zip = crear_zip_legajos(ruta_destino.get())

    #mensaje final
    mensaje_final = (
        "PROCESO FINALIZADO\n\n" + "\n\n".join(resumen)
        + f"\n\n ZIP generado correctamente en:\n{ruta_zip}"
    )

    #-----------------------
        
    messagebox.showinfo(
        "Legajos generados",
        mensaje_final
    )

def extraer_dni_nombre(nombre_carpeta):
    if " - " not in nombre_carpeta:
        return None, None
    
    partes = nombre_carpeta.split(" - ", 1)
    dni = partes[0].strip()
    nombre = partes[1].strip()

    if not dni.isdigit():
        return None, None
    
    return dni, nombre

def clasificar_archivos(carpeta):
    pdfs = []
    convertibles = []
    ignorados = []

    for nombre in os.listdir(carpeta):
        if nombre.startswith("~$"):
            ignorados.append(nombre)
            continue

        ruta = os.path.join(carpeta, nombre)

        if not os.path.isfile(ruta):
            continue

        nombre_lower = nombre.lower()
        
        if nombre_lower.endswith(EXTENSION_PDF):
            pdfs.append(nombre)
        elif nombre_lower.endswith(EXTENSIONES_WORD + EXTENSIONES_EXCEL + EXTENSIONES_IMG):
            convertibles.append(nombre)
        else:
            ignorados.append(nombre)
    return pdfs, convertibles, ignorados

def crear_carpeta_temporal(carpeta_colaborador):
    carpeta_temp = os.path.join(carpeta_colaborador, "temp_pdfs")
    os.makedirs(carpeta_temp, exist_ok=True)
    return carpeta_temp

def copiar_archivos(carpeta_origen, archivos, carpeta_temp):
    for nombre in archivos:
        origen = os.path.join(carpeta_origen, nombre)
        destino = os.path.join(carpeta_temp, nombre)
        shutil.copy2(origen, destino)

def convertir_a_pdf_guardar(ruta_origen, carpeta_temp):
    
    ruta_origen = os.path.abspath(os.path.normpath(ruta_origen))
    carpeta_temp = os.path.abspath(os.path.normpath(carpeta_temp)) 

    # Asegurar carpeta temporal
    os.makedirs(carpeta_temp, exist_ok=True)

    nombre_base = os.path.splitext(os.path.basename(ruta_origen))[0]
    salida_pdf = os.path.join(carpeta_temp, f"{nombre_base}.pdf")

    if ruta_origen.lower().endswith(EXTENSIONES_IMG):
        imagen = Image.open(ruta_origen)
        if imagen.mode != "RGB":
            imagen = imagen.convert("RGB")
        imagen.save(salida_pdf)

    elif ruta_origen.lower().endswith(EXTENSIONES_WORD):
        word = comtypes.client.CreateObject("Word.Application")
        word.Visible = False
        
        try:
            doc = word.Documents.Open(
                ruta_origen,
                ReadOnly=True)
            doc.ExportAsFixedFormat(
                OutputFileName=salida_pdf,
                ExportFormat=17)
            doc.Close(False)
        except Exception as e:
            print(f"No se pudo convertir Word: {ruta_origen}")
            print(e)
        finally:
            word.Quit()
    
    elif ruta_origen.lower().endswith(EXTENSIONES_EXCEL):
        excel = comtypes.client.CreateObject("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        try:
            wb = excel.Workbooks.Open(
                ruta_origen,
                ReadOnly=True
            )
            wb.ExportAsFixedFormat(0, salida_pdf)
            wb.Close(False)
        finally:
            excel.Quit()

    return salida_pdf

def nuevo_proceso():
    ruta_origen.set("")
    ruta_destino.set("")
    progress["value"] = 0

def cancelar():
    root.destroy()


#---------VENTANA PRINCIPAL UI----------

#VARIABLES
FUENTE = ("Segoe UI", 10)
FUENTE_BOTON = ("Segoe UI", 10, "bold")
FUENTE_TITULO = ("Segoe UI", 12, "bold")
COLOR_TITULO = "#0D1685"
COLOR_BOTON = "#6F7AA6"
COLOR_ACCION = "#1F89CC"
COLOR_ACCIONGENERADOR = "#1DB93A"
COLOR_FONDO = "#FDFEFF"

#------------------------------------

root = tk.Tk()
root.title("Generador de Legajos - HR")
root.geometry("450x370")
root.configure(bg=COLOR_FONDO)
root.resizable(False, False)  

ruta_origen = tk.StringVar()
ruta_destino = tk.StringVar()

# ----------UI ---------------------------------------

#----Titulo
tk.Label(
    root, 
    text="GENERADOR DE LEGAJOS",
    bg=COLOR_FONDO, 
    font=FUENTE_TITULO, 
    fg=COLOR_TITULO
).pack(pady=15)

#-----Frame Origen
frame_origen = tk.Frame(root, bg=COLOR_FONDO)
frame_origen.pack(fill="x", padx=20, pady=5)

tk.Label(
    frame_origen, 
    text="Carpeta de Colaboradores",
    bg=COLOR_FONDO, 
    font=FUENTE).grid(row=0, column=0, sticky="w")
tk.Entry(
    frame_origen, 
    textvariable=ruta_origen, 
    width=50,
    bg="white").grid(row=1, column=0, padx=(0,5))
tk.Button(
    frame_origen, 
    text="Seleccionar",
    width=10, 
    bg=COLOR_BOTON,
    fg="white",
    relief="flat",
    command=seleccionar_origen).grid(row=1, column=1)

#-----Frame destino
frame_destino = tk.Frame(root, bg=COLOR_FONDO)
frame_destino.pack(fill="x", padx=20, pady=5)

tk.Label(
    frame_destino, 
    text="Carpeta destino (ZIP)", 
    bg=COLOR_FONDO,
    font=FUENTE).grid(row=0, column=0, sticky="w")
tk.Entry(
    frame_destino, 
    textvariable=ruta_destino, 
    width=50,
    bg="white"). grid(row=1, column=0, padx=(0, 5))
tk.Button(
    frame_destino, 
    text="Seleccionar",
    width=10,
    bg=COLOR_BOTON,
    fg="white",
    relief="flat", 
    command=seleccionar_destino).grid(row=1, column=1)

#----------BARRA DE PROGRESO-----------
progress = ttk.Progressbar(
    root,
    orient="horizontal",
    length=300,
    mode="determinate"
)
progress.pack(pady=15)

# ---------BOTONES NUEVO Y CANCELAR
frame_botones = tk.Frame(root, bg=COLOR_FONDO)
frame_botones.pack(pady=5)

tk.Button(
    frame_botones,
    text="Nuevo Proceso",
    relief="flat",
    width=15,
    bg=COLOR_ACCION,
    fg="white",
    font=FUENTE_BOTON,
    command=nuevo_proceso
).grid(row=0, column=0, pady=5)

tk.Button(
    frame_botones,
    text="Cancelar",
    relief="flat",
    width=15,
    bg=COLOR_ACCION,
    fg="white",
    font=FUENTE_BOTON,
    command=cancelar
).grid(row=2, column=0, pady=5)

#-----Boton principal
boton_generar = tk.Button(
    root,
    text="Generar Legajos",
    relief="flat",
    bg=COLOR_ACCIONGENERADOR,
    fg="white",
    width=15,
    font=FUENTE_BOTON,
    command=generar_legajos
)
boton_generar.pack(pady=5)

# ----- Ejecucion Loop principal -----
root.mainloop()