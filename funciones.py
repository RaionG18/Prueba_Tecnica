# Librerias
from datetime import datetime, timedelta
from io import BytesIO
import pandas as pd
import requests
import time
import unidecode
import winreg
import os
import locale

from openpyxl import load_workbook
from openpyxl.utils.cell import coordinate_from_string as cfs
from openpyxl.utils.cell import get_column_letter as gcl
from urllib.parse import urljoin

# Funciones 
def getListSeparator():
    '''Retrieves the Windows list separator character from the registry'''
    aReg = winreg.ConnectRegistry(None, winreg.HKEY_CURRENT_USER)
    aKey = winreg.OpenKey(aReg, r"Control Panel\International")
    return winreg.QueryValueEx(aKey, "sList")[0]

def manejar_error(excepcion):
    print(f'Ocurrió un error: {excepcion}')

def convert_rng_to_df(tlc, l_col, l_row, sheet):
    first_col = cfs(tlc)[0]
    first_row = cfs(tlc)[1]
    rng = f"{first_col}{first_row+1}:{l_col}{l_row}"

    data_rows = []
    for row in sheet[rng]:
        data_rows.append([cell.value for cell in row])
    
    return pd.DataFrame(data_rows[2:], columns=data_rows[0])

def generar_dataframe(fecha_inicio, fecha_final):
    fecha_inicio = datetime.strptime(fecha_inicio, '%Y-%m-%d %H:%M:%S')
    fecha_final = datetime.strptime(fecha_final, '%Y-%m-%d %H:%M:%S')

    delta = timedelta(hours=1)
    fechas_intermedias = [fecha_inicio + i * delta for i in range(int((fecha_final - fecha_inicio) / delta)+1)]

    return pd.DataFrame({'fecha_hora': fechas_intermedias})

def descargar_excel(url_descarga, url_dir):
    response = requests.get(url_descarga)
    response.raise_for_status()

    with open(url_dir, 'wb') as file:
        file.write(response.content)
        
    print(f'Archivo "{url_dir}" guardado exitosamente.')
    print("----------------------------------\n")

def extraer_tablas(sheet, fecha_actual):
    # Variables
    banda_mapping = {
                     "DEMANDA MINIMA": (("00:00:00", "05:00:00"), ("22:00:00", "23:00:00")),
                     "DEMANDA MEDIA": (("06:00:00", "17:00:00"),),
                     "DEMANDA MAXIMA": (("18:00:00", "21:00:00"),)
                    }
    section_headers = ['DEMANDA MÍNIMA', 'DEMANDA MEDIA', 'DEMANDA MÁXIMA']

    # Variables temporales
    last_col = ''
    last_row = ''
    df_dict = {}  # Dictionary to hold the dataframes
    
    fecha_actual_normalized = fecha_actual.date()

    for cell in sheet['A']:  # Looping Column A only
            if cell.value in section_headers:
                tblname = cell.value  # Header of the Data Set found
                tlc = cell.coordinate  # Top Left Cell of the range
                start_row = cfs(tlc)[1]  #
                for x in range(1, sheet.max_column+1):  # Find the last used column for the data in this section
                     if cell.offset(row=1, column=x).value is None:
                        last_col = gcl(x)
                        break
                for y in range(1, sheet.max_row):  # Find the last used row for the data in this section
                    if cell.offset(row=y, column=1).value is None:
                        last_row = (start_row + y) - 1
                        if last_row != (start_row + 1):  
                            break

                # print(f"Range to convert for '{tblname}' is: '{tlc}:{last_col}{last_row}'")
                df_dict[tblname] = convert_rng_to_df(tlc, last_col, last_row, sheet)  # Convert to dataframe
                df_dict[tblname]["Planta Generadora"] = df_dict[tblname]["Planta Generadora"].apply(unidecode.unidecode)
                df_dict[tblname] = df_dict[tblname][df_dict[tblname]["Nemo"] == "JEN-C"]
                df_dict[tblname]["Banda"] = unidecode.unidecode(tblname)

                tblname_normalized = unidecode.unidecode(tblname)

                # Crear un DataFrame vacío
                df_resultante = pd.DataFrame()
                if tblname_normalized in banda_mapping:
                    horas_intervalos = banda_mapping[tblname_normalized]

                    for hora_inicio, hora_final in horas_intervalos:
                        # Generar el DataFrame
                        df_temporal = generar_dataframe(f'{fecha_actual_normalized} {hora_inicio}', f'{fecha_actual_normalized} {hora_final}')
                        df_resultante = pd.concat([df_resultante, df_temporal], ignore_index=True)

                    df_resultante["Banda"] = tblname_normalized
                    df_dict[tblname] = pd.merge(df_dict[tblname], df_resultante, on='Banda')
                    
    return df_dict
