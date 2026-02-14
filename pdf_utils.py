import os
from PyPDF2 import PdfMerger

def unir_pdfs(carpeta_pdfs, salida_pdf):
    merger = PdfMerger()

    archivos = sorted([
        f for f in os.listdir(carpeta_pdfs)
        if f.lower().endswith(".pdf")
    ])

    for pdf in archivos:
        ruta_pdf = os.path.join(carpeta_pdfs, pdf)
        merger.append(ruta_pdf)

    merger.write(salida_pdf)
    merger.close()

    return salida_pdf