import customtkinter as ctk
from tkinter import filedialog
import pandas as pd
import os
import re
import unicodedata
from docxtpl import DocxTemplate
from difflib import get_close_matches

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

excel_path = ""
carpeta_plantillas = ""

def remover_acentos(texto):
    texto = unicodedata.normalize('NFD', texto)
    return "".join([c for c in texto if unicodedata.category(c) != 'Mn'])

def limpiar_texto(valor):
    if pd.isna(valor):
        return ""
    return str(valor).strip()

def limpiar_xml(valor):
    if not valor:
        return ""
    valor = str(valor)
    valor = valor.replace("&", "y").replace("<", "").replace(">", "").replace('"', "").replace("'", "")
    return valor.strip()

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

def cargar_excel():
    global excel_path
    excel_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    label_excel.configure(text=os.path.basename(excel_path))

def cargar_plantillas():
    global carpeta_plantillas
    carpeta_plantillas = filedialog.askdirectory()
    label_plantillas.configure(text="Carpeta cargada")

def generar():
    if not excel_path or not carpeta_plantillas:
        status_label.configure(text="Carga archivos primero", text_color="red")
        return

    try:
        df = pd.read_excel(excel_path)
        df.columns = df.columns.str.strip().str.lower()

        plantillas = [f for f in os.listdir(carpeta_plantillas) if f.lower().endswith(".docx")]

        plantillas_dict = {
            normalizar(os.path.splitext(p)[0]): p for p in plantillas
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

        carpeta_salida = filedialog.askdirectory()

        generados = 0
        sin_plantilla = 0

        for idx, fila in df.iterrows():

            descripcion_raw = limpiar_texto(fila.get("descripcion", ""))
            descripcion = normalizar(descripcion_raw)

            if not descripcion:
                sin_plantilla += 1
                continue

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
                if "inefectiv" in descripcion:
                    clave_busqueda = "cambio de nombre inefectivo"
                else:
                    clave_busqueda = "cambio de nombre efectivo"

            elif "taponamiento" in descripcion:
                if "inefectiv" in descripcion:
                    clave_busqueda = "taponamiento inefectivo"
                else:
                    clave_busqueda = "taponamiento efectivo"

            elif any(x in descripcion for x in ["independizacion", "nueva conexion", "acometida", "vinculacion"]):
                if "inefectiv" in descripcion:
                    clave_busqueda = "nueva acometida inefectiva"
                else:
                    clave_busqueda = "nueva acometida efectiva espera"

            elif "revision" in descripcion and "interna" in descripcion:
                clave_busqueda = "informacion visita"

            elif "olivos" in descripcion:
                clave_busqueda = "los olivos"

            elif "red assit" in descripcion:
                clave_busqueda = "red assit"

            nombre_plantilla = None

            if clave_busqueda:
                nombre_plantilla = buscar_plantilla(clave_busqueda, plantillas_dict)

            if not nombre_plantilla:
                coincidencia = get_close_matches(descripcion, lista_claves, n=1, cutoff=0.7)
                if coincidencia:
                    nombre_plantilla = plantillas_dict[coincidencia[0]]
                else:
                    sin_plantilla += 1
                    continue

            ruta_plantilla = os.path.join(carpeta_plantillas, nombre_plantilla)

            if not os.path.exists(ruta_plantilla):
                sin_plantilla += 1
                continue

            contexto = {}

            for col in df.columns:
                llave = rename_map.get(col, col)
                valor = fila.get(col, "")

                if isinstance(valor, float) and valor.is_integer():
                    valor = int(valor)

                if llave == "nombre_radicado":
                    valor = str(valor).upper().strip()

                if llave == "control_tecnico" and valor:
                    try:
                        fecha = pd.to_datetime(valor, errors="coerce")
                        if pd.notna(fecha):
                            valor = fecha.strftime("%d/%m/%Y")
                        else:
                            valor = ""
                    except:
                        valor = ""

                if "fecha" in llave and valor:
                    try:
                        fecha = pd.to_datetime(valor, errors="coerce")
                        if pd.notna(fecha):
                            valor = fecha.strftime("%d/%m/%Y")
                        else:
                            valor = ""
                    except:
                        pass

                contexto[llave] = limpiar_xml(limpiar_texto(valor))

            try:
                doc = DocxTemplate(ruta_plantilla)
                doc.render(contexto)

                nombre = limpiar_nombre_archivo(contexto.get("nombre_radicado", f"sin_nombre_{idx}"))
                ruta_salida = os.path.join(carpeta_salida, f"oficio_{nombre}_{idx}.docx")

                doc.save(ruta_salida)
                generados += 1

            except Exception as e:
                print(f"Error: {e}")
                sin_plantilla += 1

        status_label.configure(
            text=f"Generados: {generados} | Sin plantilla: {sin_plantilla}",
            text_color="#3325AD"
        )

    except Exception as e:
        status_label.configure(text=f"Error: {str(e)}", text_color="red")

app = ctk.CTk()
app.title("Generador de Oficios Masivos")
app.geometry("620x620")

frame = ctk.CTkFrame(app, corner_radius=20)
frame.pack(padx=20, pady=20, fill="both", expand=True)

titulo = ctk.CTkLabel(frame, text="Generador de Oficios", font=("Arial", 20, "bold"))
titulo.pack(pady=15)

btn_excel = ctk.CTkButton(frame, text="Adjuntar Excel", command=cargar_excel,
                         fg_color="#A7C7E7", hover_color="#89B6E2")
btn_excel.pack(pady=10)

label_excel = ctk.CTkLabel(frame, text="Ningún archivo se ha adjuntado")
label_excel.pack()

btn_plantillas = ctk.CTkButton(frame, text="Adjuntar Plantillas", command=cargar_plantillas,
                              fg_color="#A7C7E7", hover_color="#89B6E2")
btn_plantillas.pack(pady=10)

label_plantillas = ctk.CTkLabel(frame, text="Ninguna plantilla seleccionada")
label_plantillas.pack()

btn_generar = ctk.CTkButton(frame, text="Generar Oficios",
                           command=generar,
                           fg_color="#A7C7E7",
                           hover_color="#89B6E2")
btn_generar.pack(pady=20)

status_label = ctk.CTkLabel(frame, text="")
status_label.pack(pady=10)

app.mainloop()