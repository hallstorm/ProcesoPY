import pandas as pd
import os
import pandasql as ps
import xlsxwriter
from datetime import datetime

# Carpeta raiz del proyecto
root_project_folder = os.path.dirname(os.path.abspath(__file__))

combo_df = pd.read_excel(f"{root_project_folder}/combo_table.xlsx")

# Get the list of files in the `files` folder
files = os.listdir("files")

print(files)

desired_width = 1700
pd.set_option('display.width', desired_width)
pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)

# Loop over the files
for file in files:
    # If the file is an Excel file
    if file.endswith(".xlsx"):
        # Obtener la fecha y hora actual
        fecha_hora_actual = datetime.now()

        # Formatear la fecha y hora en el formato deseado: añoMesDíaHoraMinutosSegundosMicrosegundos
        formato_deseado = "%Y%m%d%H%M%S%f"
        fecha_hora_formateada = fecha_hora_actual.strftime(formato_deseado)

        # Create an Excel writer object to save the dataframes to an Excel file
        excel_file_path = os.path.join(root_project_folder, f"output_data-{fecha_hora_formateada}.xlsx")
        excel_writer = pd.ExcelWriter(excel_file_path, engine='xlsxwriter')

        excel_file = os.path.join(root_project_folder, "files", file)

        # Read the file into a Pandas DataFrame
        ventas_df = pd.read_excel(excel_file, sheet_name="Ventas")
        control_df = pd.read_excel(excel_file, sheet_name="Control")

        # Find matching rows in ventas_df and combo_df based on "Código EAN" and "id_vao"
        matching_rows = ventas_df.merge(combo_df, left_on="Código EAN", right_on="id_vao")

        # Perform the SQL query to calculate the "Cantidad de paquetes" * "cantidad"
        sql_query = "SELECT [Fecha del documento], [Registro de tiempo], [Codigo KA/OGK], [Codigo del PDV], [Razón Social], [Calle], [Numero], [Localidad], [ean], [EAN Descripción], [Nro de factura], [Cantidad de paquetes] * [cantidad] AS [Cantidad Calculada] FROM matching_rows"
        result_df = ps.sqldf(sql_query, locals())

        result_df.rename(columns={'ean': 'Código EAN'}, inplace=True)
        result_df.rename(columns={'Cantidad Calculada': 'Cantidad de paquetes'}, inplace=True)

        # Filtrar los índices de las filas en ventas_df que coinciden con los valores de la columna "ean" en combo_df
        indices_a_eliminar = ventas_df[ventas_df["Código EAN"].isin(combo_df["id_vao"])].index

        # Eliminar las filas que cumplan la condición del paso anterior
        ventas_df = ventas_df.drop(indices_a_eliminar)

        ventas_df = pd.concat([ventas_df, result_df])

        ventas_df['Código EAN'] = ventas_df['Código EAN'].astype(str)

        # Save control_df to the first sheet named "Control"
        control_df.to_excel(excel_writer, sheet_name='Control', index=False)

        # Save result_df to the second sheet named "Ventas"
        ventas_df.to_excel(excel_writer, sheet_name='Ventas', index=False)

        # Close the Excel writer
        excel_writer.close()