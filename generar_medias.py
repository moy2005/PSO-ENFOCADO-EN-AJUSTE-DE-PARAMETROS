import os
import pandas as pd
import numpy as np

def generar_tabla_resumen_desde_archivos(ruta_archivos=".", nombre_salida="EXPERIMENTO DE AJUSTE DE PARAMETROS.xlsx"):
    archivos = [f for f in os.listdir(ruta_archivos) if f.startswith("PSO_") and f.endswith(".xlsx")]
    print(f"📁 Archivos encontrados: {len(archivos)}")

    todos_metodos = set()
    problemas = []

    for archivo in archivos:
        ruta_completa = os.path.join(ruta_archivos, archivo)
        try:
            xls = pd.ExcelFile(ruta_completa)
            hojas = xls.sheet_names
            todos_metodos.update(hojas)
            problemas.append(archivo.replace("PSO_", "").replace(".xlsx", ""))
        except Exception as e:
            print(f"⚠️ No se pudo leer el archivo {archivo}: {e}")

    todos_metodos = sorted(todos_metodos)
    resumen_datos = []

    for archivo in archivos:
        ruta_completa = os.path.join(ruta_archivos, archivo)
        nombre_problema = archivo.replace("PSO_", "").replace(".xlsx", "")
        fila = [nombre_problema]

        for metodo in todos_metodos:
            try:
                df = pd.read_excel(ruta_completa, sheet_name=metodo)

                # ⚠️ Aquí ajustamos las filas correctas según tu estructura
                factibles = df.iloc[-1, 1:-1]  # Fila de "Factibles"
                fitness_final_gen = df.iloc[-4, 1:-1]  # Última generación de valores válidos

                valores_factibles = [fitness_final_gen[i] for i in range(len(factibles)) if factibles[i] == 1]
                
                print(f"Corrida: {len(todos_metodos)})")

                if len(valores_factibles) == 0:
                    fila.append("-")
                else:
                    promedio = np.mean(valores_factibles)
                    fila.append(f"{promedio:.2e} ({len(valores_factibles)})")

            except Exception as e:
                print(f"⚠️ Error leyendo {archivo} | Hoja {metodo}: {e}")
                fila.append("-")

        resumen_datos.append(fila)

    columnas = ["Problema"] + todos_metodos
    df_resumen = pd.DataFrame(resumen_datos, columns=columnas)
    df_resumen.to_excel(nombre_salida, index=False)
    print(f"\n✅ Archivo resumen generado: {nombre_salida}")


if __name__ == "__main__":
    generar_tabla_resumen_desde_archivos(
        ruta_archivos="./",  # o "./resultados" si están en una subcarpeta
        nombre_salida="Resumen_final_PSO.xlsx"
    )
