import os
import shutil
import random
import datetime 
from datetime import timedelta 
import pandas as pd
from docx import Document
from docx2pdf import convert
from PyPDF2 import PdfReader, PdfWriter, PdfMerger
import fitz  # PyMuPDF

# --- FUNCIONES AUXILIARES ---
MESES_INGLES = ["", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

def formatear_fecha_hora(dt):
    dia = dt.day
    mes = MESES_INGLES[dt.month]
    anio = dt.year
    hora = dt.strftime("%I:%M %p").lstrip("0")
    return f"{dia}-{mes}-{anio} {hora} EST"

def generar_tiempos(fecha_firma_str):
    try:
        dt_base = datetime.datetime.strptime(fecha_firma_str, "%m/%d/%Y")
    except ValueError:
        return None
    
    hora_inicio = random.randint(14, 17)
    minuto_inicio = random.randint(0, 59)
    dt1 = dt_base.replace(hour=hora_inicio, minute=minuto_inicio)
    dt2 = dt1 + timedelta(minutes=20)
    dt3 = dt2 + timedelta(minutes=3)
    return {
        "Process started": formatear_fecha_hora(dt1),
        "Document viewed": formatear_fecha_hora(dt2),
        "Document accepted & signed": formatear_fecha_hora(dt3),
        "Document has been completed": formatear_fecha_hora(dt3)
    }

def concatenar_direccion(row):
    valores = [str(v) for v in row[['address1', 'address2', 'others']] if pd.notna(v)]
    return ', '.join(valores)

def process_medical_requests():
    root_dir = os.getcwd()
    # Configuración de archivos basada en detección automática
    archivos = os.listdir(root_dir)
    nombre_fa = next((f for f in archivos if f.lower().endswith('.pdf')), None)
    nombre_word = next((f for f in archivos if f.lower().endswith('.docx')), None)

    if not nombre_fa:
        print("❌ Error: No se detectó archivo PDF (HIPAA firmado).")
        return

    # Interfaz de entrada (simplificada para script)
    print("=== SISTEMA DE SOLICITUDES MÉDICAS AUTOMATIZADO ===")
    p1 = int(input("Primera página a extraer: "))
    p2 = int(input("Segunda página a extraer: "))
    client_name = input("Nombre del cliente: ")
    f_inicio = input("Fecha INICIO (mm/dd/yyyy): ")
    f_solicitud = input("Fecha FINAL (mm/dd/yyyy): ")
    f_firma = input("Fecha de FIRMA (mm/dd/yyyy): ")
    
    # Lógica de creación de carpetas y edición de HIPAA (fitz)
    # ... (Se mantiene la lógica de extracción y edición del PDF maestro)
    #

    print("\n✨ Proceso completado exitosamente.")

if __name__ == "__main__":
    process_medical_requests()