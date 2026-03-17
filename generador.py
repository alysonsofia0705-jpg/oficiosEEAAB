import pandas as pd
from docxtpl import DocxTemplate

# leer excel
df = pd.read_excel("base.xlsx")

print(df.columns)

for i, fila in df.iterrows():

    doc = DocxTemplate("cambio de nombre efectivo.docx")

    contexto = {
        "nombre": fila["nombre"],
        "apellido": fila["apellido"],
        "direccion": fila["direccion"],
        "telefono": fila["telefono"],
        "correo": fila["correo"],
        "descripcion": fila["descripcion"],
        "medidor": fila["medidor"],
        "contrato": fila["Cta.contrato"],
        "radicado": fila["nombre.radicado"],
        "fecha_radicado": fila["fecha.de.radicado"]
    }

    doc.render(contexto)

    nombre_archivo = f"oficio_{fila['nombre']}_{fila['apellido']}.docx"

    doc.save(nombre_archivo)

print("Oficios generados correctamente")