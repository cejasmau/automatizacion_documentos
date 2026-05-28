"""
Generador de actas a partir de Excel y plantilla Word.
Uso:
    python generador.py [--excel RUTA] [--plantilla RUTA] [--salida DIR]
"""

# %%
# Importación de librerías
import pandas as pd
from docxtpl import DocxTemplate
import os
import sys
import argparse

MESES = {
    1: "enero", 2: "febrero", 3: "marzo", 4: "abril",
    5: "mayo", 6: "junio", 7: "julio", 8: "agosto",
    9: "septiembre", 10: "octubre", 11: "noviembre", 12: "diciembre"
}

# Funciones auxiliares

def obtener_autoridad(row):
    """
    Determina la autoridad vigente y su DNI para una fila, según la fecha de entrega.
    
    Lógica:
    - Extrae el año de la fecha_entrega.
    - Obtiene la fecha de vencimiento del mandato para ese año (columna aut_{año}_venc).
    - Si la fecha de vencimiento es válida y ya pasó (es anterior a fecha_entrega), 
      se toma la autoridad del año siguiente; en caso contrario, la del año actual.
    
    Retorna una Series con:
        "autoridad": nombre de la autoridad (aut_{año})
        "autoridad_dni": DNI de la autoridad (aut_{año}_dni)
    """
    anio = row["fecha_entrega"].year
    
    venc = row[f"aut_{anio}_venc"]

    venc_dt = pd.to_datetime(venc, errors = 'coerce')
    
    if pd.notna(venc_dt) and venc_dt < row["fecha_entrega"]:
        anio_autoridad = anio + 1
    else:
        anio_autoridad = anio
    
    return pd.Series({
        "autoridad": row[f"aut_{anio_autoridad}"],
        "autoridad_dni": row[f"aut_{anio_autoridad}_dni"]
    })

def obtener_domicilio(row):
    """
    Obtiene el domicilio correspondiente al año de la fecha de entrega.
    
    Retorna el valor de la columna dom_{año} para esa fila.
    """
    anio = row["fecha_entrega"].year
    return row[f"dom_{anio}"]

# Función principal

def generar_actas(excel_path, plantilla_path, output_dir):
    """
    Genera documentos Word a partir de un archivo Excel.

    Parámetros:
        excel_path    (str): ruta al archivo Excel.
        plantilla_path(str): ruta a la plantilla .docx.
        output_dir    (str): carpeta donde se guardarán las actas generadas.
    """

    # 1. Leer el Excel 
    df = pd.read_excel(excel_path, sheet_name = "Hoja1")
    df_ent = pd.read_excel(excel_path, sheet_name = "Hoja2")

    # 2. Chequear tipos de datos
    df_ent["cuit"] = df_ent["cuit"].astype("Int64")
    df_ent["matricula"] = df_ent["matricula"].astype("Int64")

    # 3. Seleccionar columnas de entidades (años 2022 a 2025 en este ejemplo)
    columnas_entidad = ["entidad_adm", "cuit", "matricula"]
    for anio in range(2022, 2026):
        columnas_entidad += [
            f"aut_{anio}", f"aut_{anio}_dni", f"aut_{anio}_venc", f"dom_{anio}"
        ]

    # 4. Combinar DataFrames
    df = df.merge(df_ent[columnas_entidad], 
                  on = "entidad_adm", 
                  how = "left")

    # 5. Aplicar funciones fila por fila
    df[["autoridad", "autoridad_dni"]] = df.apply(obtener_autoridad, axis = 1)
    df["domicilio"] = df.apply(obtener_domicilio, axis = 1)


    # 6. Cargar la plantilla
    plantilla = DocxTemplate(plantilla_path)

    # 7. Crear la carpeta de salida si no existe
    os.makedirs(output_dir, exist_ok=True)

  # 8. Generar un acta por cada id_acta
    for id_acta, grupo in df.groupby("id_acta"):
        # Datos comunes (primera fila del grupo)
        datos_comunes = grupo.iloc[0].to_dict()
        fecha_entrega = datos_comunes["fecha_entrega"]

        # Datos de la tabla (elementos de cada fila)
        elementos = grupo[["elemento_acta", "cantidad", "unidad_prod"]].to_dict(orient="records")

        # Contexto para la plantilla
        contexto = {
            **datos_comunes,
            "elementos": elementos,
            "dia": fecha_entrega.day,
            "mes": MESES[fecha_entrega.month],
            "anio": fecha_entrega.year
        }

        # Renderizar y guardar
        plantilla.render(contexto)
        nombre_archivo = os.path.join(output_dir, f"acta_{id_acta}.docx")
        plantilla.save(nombre_archivo)

    print(f"✅ Actas generadas correctamente en: {output_dir}")

# Punto de entrada del script

def main():
    parser = argparse.ArgumentParser(description = "Generador de actas desde Excel")
    parser.add_argument("--excel", default="datos/prueba.xlsx")
    parser.add_argument("--plantilla", default="plantilla.docx")
    parser.add_argument("--salida", default="actas")
    args = parser.parse_args()

    # Detectar si estamos corriendo como script o como ejecutable
    if getattr(sys, 'frozen', False):
        # Ejecutable: usar la carpeta donde está el .exe
        base_dir = os.path.dirname(sys.executable)
    else:
        # Script: usar la carpeta del script
        base_dir = os.path.dirname(os.path.abspath(__file__))

    excel_abs = os.path.normpath(os.path.join(base_dir, args.excel))
    plantilla_abs = os.path.normpath(os.path.join(base_dir, args.plantilla))
    salida_abs = os.path.normpath(os.path.join(base_dir, args.salida))

    generar_actas(excel_abs, plantilla_abs, salida_abs)

if __name__ == "__main__":
    main()





