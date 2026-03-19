import os
import re
import pandas as pd
from docxtpl import DocxTemplate

# FUNCIONES

def limpiar_texto(valor):
    if pd.isna(valor):
        return ""
    return str(valor).strip()

def normalizar(texto):
    return limpiar_texto(texto).lower()

def limpiar_nombre_archivo(valor):
    texto = limpiar_texto(valor)
    texto = re.sub(r'[\\/*?:"<>|]', "", texto)
    texto = texto.replace("  ", " ").strip()
    return texto

# LEER EXCEL

df = pd.read_excel("base.xlsx")

print("\nColumnas detectadas:")
print(df.columns)

# DETECTAR PLANTILLAS

carpeta_plantillas = "Formatos_Cruce"

plantillas = {}

for archivo in os.listdir(carpeta_plantillas):

    if archivo.endswith(".docx"):

        nombre_limpio = normalizar(archivo.replace(".docx", ""))

        plantillas[nombre_limpio] = archivo


print("\nPlantillas detectadas:")

for p in plantillas:
    print("-", p)


# RELACIONA DESCRIPCION PLANTILLA

mapa_descripciones = {
    "taponamiento efectivo": "Taponamiento Efectivo.docx",
    "taponamiento inefectivo": "Taponamiento Inefectivo Deuda.docx",
    "activar factura virtual": "ACTIVAR FACTURA VIRTUAL (1).docx",
    "desactivar factura virtual": "Desactivar factura virtual.docx",
    "cambio de nombre efectivo" : "cambio de nombre efectivo.docx",
    "cambio de nombre inefectivo": "cambio de nombre inefectivo.docx",
    "independdizacion inefectiva" : "independizacion inefectiva deuda.docx",
    "independizacon efectiva" : "NUEVA ACOMETIDA EFECTIVA ESPERA.docx",
    "nueva acometida inefectiva" : "nueva acometida inefectiva documentos.docx",
    "nueva acometida efectiva" : "NUEVA ACOMETIDA EFECTIVA ESPERA.docx",
    "visita": "Informacion visita (1).docx",
    "normalizacion efectivaa" : "Normalizacion Proximo A Ejecutar.docx",
    "Paz y salvo efectivo": "Paz y salvo efectivo.docx",
    "Paz y salvo efectivo": "Paz y salvo.docx",
    "Paz y salvo inefectivo" : "Paz y salvo inefectivo.docx",
    "reconexion efectiva" : "reconexion efectiva.docx",
    "reconexion inefectiva" : "Reconexion inefectiva Deuda.docx",
    "Revision con geofono efectiva" : "REVISION CON GEOFONO EFECTIVA (1).docx",
    "Revision con geofono inefectiva" : "REVISION CON GEOFONO INEFECTIVA (1).docx",
    "suspension efectiva" : "Suspension efectiva.docx",
    "suspension inefectiva" : "Suspension inefectiva deuda.docx",
    "vinculacion inefectiva" : "vincu inefectiva sin putos hidraulicos.docx",
}

  

# RENOMBRE COLUMNAS

rename_map = {
    "Cta.contrato": "cta_contrato",
    "Interl.comercial": "interl_comercial",
    "control.tecnico": "control_tecnico",
    "calle.2": "calle_2",
    "cuenta.interna": "cuenta_interna",
    "nombre.radicado": "nombre_radicado",
    "fecha.de.radicado": "fecha_de_radicado"
}


# GENERAR OFICIOS


for idx, fila in df.iterrows():

    descripcion = normalizar(fila.get("descripcion", ""))

    nombre_plantilla = mapa_descripciones.get(descripcion)

    if not nombre_plantilla:

        print("Fila", idx, "-> no se encontró plantilla para:", descripcion)

        continue


    ruta_plantilla = os.path.join(carpeta_plantillas, nombre_plantilla)

    if not os.path.exists(ruta_plantilla):

        print("Fila", idx, "-> plantilla no encontrada:", nombre_plantilla)

        continue


    # construir contexto

    contexto = {}

    for col in df.columns:

        llave = rename_map.get(col, col)

        contexto[llave] = limpiar_texto(fila.get(col, ""))


    doc = DocxTemplate(ruta_plantilla)

    doc.render(contexto)


    nombre_limpio = limpiar_nombre_archivo(contexto.get("nombre", ""))

    apellido_limpio = limpiar_nombre_archivo(contexto.get("apellido", ""))


    nombre_archivo = f"oficio_{nombre_limpio}_{apellido_limpio}_{idx}.docx"


    doc.save(nombre_archivo)


    print("Generado:", nombre_archivo)


print("\nOficios generados correctamente")