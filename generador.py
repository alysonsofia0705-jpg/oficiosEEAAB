import os
import re
import sys
import pandas as pd
from docxtpl import DocxTemplate
from difflib import get_close_matches


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
    return texto.strip()

def limpiar_xml(valor):
    if not valor:
        return ""
    valor = str(valor)
    valor = valor.replace("&", "y")
    valor = valor.replace("<", "")
    valor = valor.replace(">", "")
    valor = valor.replace('"', "")
    valor = valor.replace("'", "")
    return valor.strip()

def ruta_base():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


base_path = ruta_base()

ruta_excel = os.path.join(base_path, "base.xlsx")
df = pd.read_excel(ruta_excel)
df.columns = df.columns.str.strip().str.lower()

print("\nColumnas detectadas:")
print(df.columns.tolist())

for col in df.columns:
    if "fecha" in col or "control" in col:
        df[col] = pd.to_datetime(df[col], errors='coerce').dt.strftime("%d/%m/%Y").fillna("")

carpeta_plantillas = os.path.join(base_path, "Formatos_Cruce")

plantillas = [f for f in os.listdir(carpeta_plantillas) if f.endswith(".docx")]

print("\nPlantillas disponibles:")
for p in plantillas:
    print(p)

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

    print(f"\nFila {idx} → descripción: {descripcion}")

    if not descripcion:
        print(f"Fila {idx}: descripción vacía")
        sin_plantilla += 1
        continue

    nombre_plantilla = None

    if any(p in descripcion for p in ["vinculacion", "vinculación"]):
        for clave, archivo in plantillas_dict.items():
            if "nueva acometida efectiva espera" in clave:
                nombre_plantilla = archivo
                break

        if not nombre_plantilla:
            print("No se encontró la plantilla de vinculación")
            sin_plantilla += 1
            continue

    elif "independizacion" in descripcion:
        for clave, archivo in plantillas_dict.items():
            if "nueva acometida efectiva espera" in clave:
                nombre_plantilla = archivo
                break

    elif "revisiones internas" in descripcion:
        for clave, archivo in plantillas_dict.items():
            if "informacion visita" in clave:
                nombre_plantilla = archivo
                break
    
    elif "revision interna" in descripcion:
        for clave, archivo in plantillas_dict.items():
            if "informacion visita" in clave:
                nombre_plantilla = archivo
                break

    if not nombre_plantilla:
        coincidencia = get_close_matches(descripcion, lista_claves, n=1, cutoff=0.4)

        if not coincidencia:
            print(f"Fila {idx}: no se encontró plantilla para '{descripcion_raw}'")
            sin_plantilla += 1
            continue

        nombre_plantilla = plantillas_dict[coincidencia[0]]

    ruta_plantilla = os.path.join(carpeta_plantillas, nombre_plantilla)

    if not os.path.exists(ruta_plantilla):
        print(f"Fila {idx}: plantilla no existe {nombre_plantilla}")
        sin_plantilla += 1
        continue

    contexto = {}

    for col in df.columns:
        llave = rename_map.get(col, col)
        valor = fila.get(col, "")

        if isinstance(valor, float) and valor.is_integer():
            valor = int(valor)

        contexto[llave] = limpiar_xml(limpiar_texto(valor))

    try:
        doc = DocxTemplate(ruta_plantilla)
        doc.render(contexto)

        nombre_limpio = limpiar_nombre_archivo(contexto.get("nombre", "sin_nombre"))
        apellido_limpio = limpiar_nombre_archivo(contexto.get("apellido", "sin_apellido"))

        nombre_archivo = os.path.join(
            base_path,
            f"oficio_{nombre_limpio}_{apellido_limpio}_{idx}.docx"
        )

        doc.save(nombre_archivo)

        print(f"✓ Generado: {nombre_archivo}")
        generados += 1

    except Exception as e:
        print(f"Fila {idx}: error — {e}")

print("\nRESUMEN:")
print(f"Oficios generados: {generados}")
print(f"Sin plantilla: {sin_plantilla}")