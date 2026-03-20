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
    texto = limpiar_texto(texto).lower()
    texto = re.sub(r'[^\w\s]', '', texto)
    return texto

def limpiar_nombre_archivo(valor):
    texto = limpiar_texto(valor)
    texto = re.sub(r'[\\/*?:"<>|]', "", texto)
    texto = texto.replace("  ", " ").strip()
    return texto

# LEER EXCEL
df = pd.read_excel("base.xlsx")
df.columns = df.columns.str.strip().str.lower()

print("\nColumnas detectadas:")
print(df.columns.tolist())

# CONFIGURACIÓN
carpeta_plantillas = "Formatos_Cruce"

mapa_descripciones = {
    "taponamiento efectivo": "Taponamiento Efectivo.docx",
    "taponamiento inefectivo": "Taponamiento Inefectivo Deuda.docx",
    "activar factura virtual": "ACTIVAR FACTURA VIRTUAL (1).docx",
    "desactivar factura virtual": "Desactivar factura virtual.docx",
    "cambio de nombre efectivo": "cambio de nombre efectivo.docx",
    "cambio de nombre inefectivo": "cambio de nombre inefectivo.docx",
    "independdizacion inefectiva": "independizacion inefectiva deuda.docx",
    "independizacon efectiva": "NUEVA ACOMETIDA EFECTIVA ESPERA.docx",
    "nueva acometida inefectiva": "nueva acometida inefectiva documentos.docx",
    "nueva acometida efectiva": "NUEVA ACOMETIDA EFECTIVA ESPERA.docx",
    "visita": "Informacion visita (1).docx",
    "normalizacion efectivaa": "Normalizacion Proximo A Ejecutar.docx",
    "paz y salvo efectivo": "Paz y salvo efectivo.docx",
    "paz y salvo inefectivo": "Paz y salvo inefectivo.docx",
    "reconexion efectiva": "reconexion efectiva.docx",
    "reconexion inefectiva": "Reconexion inefectiva Deuda.docx",
    "revision con geofono efectiva": "REVISION CON GEOFONO EFECTIVA (1).docx",
    "revision con geofono inefectiva": "REVISION CON GEOFONO INEFECTIVA (1).docx",
    "suspension efectiva": "Suspension efectiva.docx",
    "suspension inefectiva": "Suspension inefectiva deuda.docx",
    "vinculacion inefectiva": "vincu inefectiva sin putos hidraulicos.docx",
    "revision interna": "informacion visita (1).docx",
}

rename_map = {
    "cta.contrato": "cta_contrato",
    "interl.comercial": "interl_comercial",
    "control.tecnico": "control_tecnico",
    "calle.2": "calle_2",
    "cuenta.interna": "cuenta_interna",
    "nombre.radicado": "nombre_radicado",
    "fecha.de.radicado": "fecha_de_radicado",
    "direccion": "direccion",
    "correo": "correo",
    "telefono": "telefono",
    "nombre": "nombre",
    "apellido": "apellido",
    "entrada": "entrada",
}

# GENERAR OFICIOS
for idx, fila in df.iterrows():

    descripcion_raw = limpiar_texto(fila.get("descripcion", ""))
    descripcion = normalizar(descripcion_raw)

    if not descripcion:
        print(f"Fila {idx}: descripción vacía, se omite.")
        continue

    # BUSCAR PLANTILLA usando el mapa
    nombre_plantilla = None

    for clave, archivo in mapa_descripciones.items():
        if normalizar(clave) in descripcion or descripcion in normalizar(clave):
            nombre_plantilla = archivo
            break

    if not nombre_plantilla:
        print(f"Fila {idx}: no se encontró plantilla para '{descripcion_raw}'")
        continue

    ruta_plantilla = os.path.join(carpeta_plantillas, nombre_plantilla)

    if not os.path.exists(ruta_plantilla):
        print(f"Fila {idx}: archivo de plantilla no existe en disco: {nombre_plantilla}")
        continue

    # CONSTRUIR CONTEXTO
    contexto = {}

    for col in df.columns:
        llave = rename_map.get(col, col)
        valor = fila.get(col, "")

        if "fecha" in llave:
            if pd.notna(valor):
                try:
                    fecha_str = str(valor)
                    fecha_limpia = pd.to_datetime(fecha_str).strftime("%d/%m/%Y")
                    contexto[llave] = fecha_limpia
                except Exception:
                    contexto[llave] = str(valor).split(" ")[0].split("T")[0]
            else:
                contexto[llave] = ""
        else:
            contexto[llave] = limpiar_texto(valor)

    # IMPRIMIR CONTEXTO PARA VERIFICAR FECHAS
    print(f"\nFila {idx} - contexto de fechas:")
    for k, v in contexto.items():
        if "fecha" in k:
            print(f"  {k}: {v}")

    # GENERAR DOCUMENTO
    try:
        doc = DocxTemplate(ruta_plantilla)
        doc.render(contexto)

        nombre_limpio = limpiar_nombre_archivo(contexto.get("nombre", "sin_nombre"))
        apellido_limpio = limpiar_nombre_archivo(contexto.get("apellido", "sin_apellido"))

        nombre_archivo = f"oficio_{nombre_limpio}_{apellido_limpio}_{idx}.docx"
        doc.save(nombre_archivo)

        print(f"✓ Generado: {nombre_archivo}")

    except Exception as e:
        print(f"Fila {idx}: error al generar documento — {e}")

print("\nProceso finalizado.")