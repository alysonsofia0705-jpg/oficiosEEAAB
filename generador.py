import os
import re
import sys
import unicodedata
import pandas as pd
from docxtpl import DocxTemplate
from difflib import get_close_matches

def remover_acentos(texto):
    texto = unicodedata.normalize('NFD', texto)
    return "".join([c for c in texto if unicodedata.category(c) != 'Mn'])

def limpiar_texto(valor):
    if pd.isna(valor):
        return ""
    return str(valor).strip()

def normalizar(texto):
    texto = limpiar_texto(texto).lower()
    texto = remover_acentos(texto)
    texto = re.sub(r'[^\w\s]', '', texto)
    texto = re.sub(r'\s+', ' ', texto)
    return texto.strip()

def limpiar_nombre_archivo(valor):
    texto = limpiar_texto(valor)
    texto = re.sub(r'[\\/*?:"<>|]', "", texto)
    return texto.strip()

def limpiar_xml(valor):
    if not valor:
        return ""
    valor = str(valor)
    valor = valor.replace("&", "y").replace("<", "").replace(">", "").replace('"', "").replace("'", "")
    return valor.strip()

def ruta_base():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

def contiene(texto, palabra):
    return palabra in texto.replace(" ", "")

def buscar_plantilla(clave_busqueda, plantillas_dict):
    clave_norm = normalizar(clave_busqueda)
    
    if clave_norm in plantillas_dict:
        return plantillas_dict[clave_norm]
    
    for clave, archivo in plantillas_dict.items():
        if clave_norm in clave or clave in clave_norm:
            return archivo
            
    palabras_maestras = ["geofono", "suspension", "reconexion", "taponamiento", "nombre", "acometida"]
    filtro = next((p for p in palabras_maestras if p in clave_norm), None)
    
    if filtro:
        candidatos = [k for k in plantillas_dict.keys() if filtro in k]
        coincidencias = get_close_matches(clave_norm, candidatos, n=1, cutoff=0.3)
        if coincidencias:
            return plantillas_dict[coincidencias[0]]
            
    return None

base_path = ruta_base()

ruta_excel = os.path.join(base_path, "base.xlsx")
df = pd.read_excel(ruta_excel)
df.columns = df.columns.str.strip().str.lower()

carpeta_plantillas = os.path.join(base_path, "Formatos_Cruce")
plantillas = [f for f in os.listdir(carpeta_plantillas) if f.lower().endswith(".docx")]

plantillas_dict = {
    normalizar(p.replace(".docx", "")): p for p in plantillas
}

lista_claves = list(plantillas_dict.keys())

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

generados = 0
sin_plantilla = 0

for idx, fila in df.iterrows():
    
    descripcion_raw = limpiar_texto(fila.get("descripcion", ""))
    descripcion = normalizar(descripcion_raw)
    
    print(f"\nFila {idx} → {descripcion_raw}")

    if not descripcion:
        print("Sin descripción")
        sin_plantilla += 1
        continue

    nombre_plantilla = None
    clave_busqueda = None

    if "geofono" in descripcion or contiene(descripcion, "geofono"):
        if "inefectiv" in descripcion or contiene(descripcion, "inefectiva"):
            clave_busqueda = "revision con geofono inefectiva"
        else:
            clave_busqueda = "revision con geofono efectiva"

    elif "suspension" in descripcion or contiene(descripcion, "suspension"):
        if "inefectiv" in descripcion or contiene(descripcion, "inefectiva"):
            clave_busqueda = "suspension inefectiva"
        else:
            clave_busqueda = "suspension efectiva"

    elif "reconexion" in descripcion or contiene(descripcion, "reconexion"):
        if "inefectiv" in descripcion or contiene(descripcion, "inefectiva"):
            clave_busqueda = "reconexion inefectiva"
        else:
            clave_busqueda = "reconexion efectiva"

    elif "cambio de nombre" in descripcion or contiene(descripcion, "cambiodenombre"):
        clave_busqueda = "cambio de nombre inefectivo" if "inefectiv" in descripcion else "cambio de nombre efectivo"

    elif "taponamiento" in descripcion:
        clave_busqueda = "taponamiento inefectivo" if "inefectiv" in descripcion else "taponamiento efectivo"

    elif any(x in descripcion for x in ["independizacion", "nueva conexion", "acometida", "vinculacion"]):
        clave_busqueda = "nueva acometida inefectiva" if "inefectiv" in descripcion else "nueva acometida efectiva espera"

    elif "revision" in descripcion and "interna" in descripcion:
        clave_busqueda = "informacion visita"

    elif "olivos" in descripcion:
        clave_busqueda = "los olivos"

    elif "red assit" in descripcion:
        clave_busqueda = "red assit"

    if clave_busqueda:
        print(f"Clave usada: {clave_busqueda}")

        nombre_plantilla = buscar_plantilla(clave_busqueda, plantillas_dict)

    if not nombre_plantilla:
        coincidencia = get_close_matches(descripcion, lista_claves, n=1, cutoff=0.7)

        if coincidencia:
            nombre_plantilla = plantillas_dict[coincidencia[0]]
            print(f"Por similitud: {coincidencia[0]}")
        else:
            print("Sin plantilla")
            sin_plantilla += 1
            continue

    print(f"Plantilla: {nombre_plantilla}")

    ruta_plantilla = os.path.join(carpeta_plantillas, nombre_plantilla)

    if not os.path.exists(ruta_plantilla):
        print("No existe archivo")
        sin_plantilla += 1
        continue

    contexto = {}

    for col in df.columns:
        llave = rename_map.get(col, col)
        valor = fila.get(col, "")

        if isinstance(valor, float) and valor.is_integer():
            valor = int(valor)
        if "fecha" in llave and valor:
            try:
                valor = pd.to_datetime(valor).strftime("%d/%m/%Y")
            except:
                pass
        contexto[llave] = limpiar_xml(limpiar_texto(valor))

    try:
        doc = DocxTemplate(ruta_plantilla)
        doc.render(contexto)
        nombre_oficio = limpiar_nombre_archivo(contexto.get("nombre_radicado", f"sin_nombre_{idx}"))
        nombre_archivo = os.path.join(base_path, f"oficio_{nombre_oficio}_{idx}.docx")
        doc.save(nombre_archivo)
        print("Generado")
        generados += 1

    except Exception as e:
        print(f"Error: {e}")
        sin_plantilla += 1

print(f"\nOficios generados: {generados}")
print(f"Sin plantilla: {sin_plantilla}")