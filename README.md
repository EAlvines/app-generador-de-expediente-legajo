# 📁 Generador de Legajos - Aplicación Desktop en Python

Aplicación de escritorio desarrollada en **Python** para el área de **Recursos Humanos**, que automatiza la creación de legajos digitales a partir de múltiples archivos por colaborador.

Permite consolidar documentos en distintos formatos (PDF, Word, Excel, imágenes) en **un solo PDF por colaborador** y generar un **ZIP final listo para carga o envío**.

<img width="345" height="230" alt="image" src="https://github.com/user-attachments/assets/6c816515-73cf-4bd5-a070-2620184ec15b" />

---

## 🎯 Objetivo del Proyecto

- Automatizar un proceso manual y repetitivo en HR
- Reducir tiempos operativos
- Minimizar errores humanos
- Fortalecer habilidades en automatización y desarrollo de aplicaciones

Este proyecto forma parte de mi portafolio como Analista de Datos con enfoque en automatización.

---

## 🧠 Flujo de la Aplicación

1. Seleccionar carpeta origen:
   - Puede contener múltiples carpetas de colaboradores
   - O una sola carpeta individual
   - Formato esperado:
     ```
     DNI - NOMBRE APELLIDO
     ```

2. Seleccionar carpeta destino

3. Por cada colaborador:
   - Identifica PDFs originales
   - Convierte Word, Excel e imágenes a PDF
   - Centraliza todos los PDFs en carpeta temporal
   - Une todos los PDFs en un único archivo final

4. Genera:
   - `Legajo_Carga.zip`
   - Contiene únicamente los PDFs finales

---

## 📂 Estructura Esperada

### Carpeta Origen

  Carpeta_Padre/
  │
  ├── 70000000 - PEPE GUIDO/
  │ ├── contrato.docx
  │ ├── dni.png
  │ ├── documentos.pdf
  │ └── temp_pdfs/
  │
  └── 70000090 - PEPE AGUINALDO/
  ├── archivo.xlsx
  ├── foto.jpg
  └── temp_pdfs/

### Resultado Final
  
  Legajo_Carga.zip
  │
  ├── 70000000 - PEPE GUIDO.pdf
  └── 70000090 - PEPE AGUINALDO.pdf

---

## 🖥️ Interfaz

- GUI desarrollada con Tkinter
- Barra de progreso por colaborador
- Validación de carpetas
- Botones:
  - Seleccionar origen
  - Seleccionar destino
  - Generar legajos
  - Nuevo proceso
  - Cancelar

---

## 🛠️ Tecnologías Utilizadas

- Python 3.13
- Tkinter
- Pillow (manejo de imágenes)
- PyPDF2 (unión de PDFs)
- win32com (automatización Word y Excel)
- PyInstaller (generación de ejecutable .exe)

---

## ⚠️ Requisitos

- Sistema operativo Windows
- Microsoft Word y Excel instalados
- Permisos de lectura y escritura en carpetas seleccionadas

---

## 🚀 Ejecutable

El proyecto puede compilarse como archivo `.exe` utilizando PyInstaller

---

## 👩‍💻 Autora

Emi
Analista de Datos | Automatización | Python
📍 Lima, Perú
