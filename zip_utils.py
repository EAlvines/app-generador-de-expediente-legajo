import os
import zipfile

def crear_zip_legajos(carpeta_destino):
    ruta_zip = os.path.join(carpeta_destino, "Legajo_carga.zip")

    with zipfile.ZipFile(ruta_zip, "w", zipfile.ZIP_DEFLATED) as zipf:
        for archivo in os.listdir(carpeta_destino):
            if archivo.lower().endswith(".pdf"):
                ruta_pdf = os.path.join(carpeta_destino, archivo)
                zipf.write(ruta_pdf, arcname=archivo)

    return ruta_zip