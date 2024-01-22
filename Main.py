from funciones import *

# Fechas para la muestra
fecha_inicial = datetime(2023, 1, 1)
fecha_final = datetime(2023, 6, 30)


if __name__ == "__main__":
    base_url = 'https://www.amm.org.gt/pdfs2/programas_despacho/'
    dir1_url = '01_PROGRAMAS_DE_DESPACHO_DIARIO'
    dir2_url = '01_PROGRAMAS_DE_DESPACHO_DIARIO'

    locale.setlocale(locale.LC_TIME, 'es_ES.utf8')

    fecha_actual = fecha_inicial

    # Define a list to store downloaded file paths
    downloaded_files = []

    while fecha_actual <= fecha_final:
        
        fecha_str_doc = fecha_actual.strftime("%d%m%Y")
        fecha_str_dir = fecha_actual.strftime("%m_%B").upper()
        year_str = fecha_actual.strftime("%Y")
        
        url_dir = f'{dir1_url}/{year_str}/{dir2_url}/{fecha_str_dir}/WEB{fecha_str_doc}.xlsx'
        url_descarga = urljoin(base_url, url_dir)
        os.makedirs(os.path.dirname(url_dir), exist_ok=True)

        # Descargar archivo actual
        try:
            descargar_excel(url_descarga, url_dir)
            downloaded_files.append(url_dir)

        except requests.exceptions.RequestException as err:
            manejar_error(err)

        fecha_actual += timedelta(days=1)

    # Variables temporales
    df_calculos = pd.DataFrame()
    fecha_actual = fecha_inicial
    sheet_name = 'LDM'
    
    
    # Process the data after downloading all files
    for file_path in downloaded_files:
        
        fecha_str_dir = fecha_actual.strftime("%m_%B").upper()
        year_str = fecha_actual.strftime("%Y")

        # Abrir documuento actual
        wb = load_workbook(file_path)

        if sheet_name not in wb.sheetnames:
            raise ValueError(f'No se encontró la hoja con el nombre "{sheet_name}" en el archivo Excel.')

        # Seleccionar la hoja LDM
        sheet = wb[sheet_name]

        # Extraer las 3 tablas de cada archivo diario
        df_dict = extraer_tablas(sheet, fecha_actual)

        # Dataframes Temporales
        df_demandas = pd.DataFrame()
        
        # Concatenacion de todos los dataframe del dia
        for table_name, df in df_dict.items():
            df_demandas = pd.concat([df_demandas, df], ignore_index=True)
        
        # Reordenar el Dataframe
        df_demandas.columns = df.columns.str.strip()
        df_demandas = df_demandas.rename(columns={'Costo en US$/MWH': 'Costo'})
        df_demandas = df_demandas[['fecha_hora', 'Nemo', 'Planta Generadora', 'Potencia Disponible', 'Costo', 'FPNE', 'Banda']]
        df_demandas = df_demandas.sort_values(by='fecha_hora')

        # Leer archivos POE y Generacion
        df_poe = pd.read_csv('POE.csv')
        df_poe['fecha_hora'] = pd.to_datetime(df_poe['fecha_hora'])

        df_generacion = pd.read_csv('Generacion.csv')
        df_generacion['fecha_hora'] = pd.to_datetime(df_generacion['fecha_hora'])
        
        # Filtrar por la fecha específica
        fecha_especifica = fecha_actual.date()
        df_poe = df_poe[df_poe['fecha_hora'].dt.date == pd.to_datetime(fecha_especifica).date()]
        df_generacion = df_generacion[df_generacion['fecha_hora'].dt.date == pd.to_datetime(fecha_especifica).date()]

        df_demandas = pd.merge(df_demandas, df_poe, on='fecha_hora')
        df_demandas = pd.merge(df_demandas, df_generacion, on='fecha_hora')

        # Calculos de Indicador
        df_demandas['Indicador'] = (df_demandas['POE'] > df_demandas['Costo']).astype(int)

        # Calculos de Liquidaciones
        df_demandas['Liquidacion POE'] = df_demandas['POE']*df_demandas['generacion']
        df_demandas['Liquidacion CVG'] = df_demandas['Costo']*df_demandas['generacion']

        # Calculos de Agentes
        df_demandas['Agente A'] = df_demandas['Indicador']*(df_demandas['Liquidacion POE'] - df_demandas['Liquidacion CVG'])
        df_demandas['Agente B'] = (df_demandas['Indicador']-1)*(df_demandas['Liquidacion POE'] - df_demandas['Liquidacion CVG'])

        nombre_archivo_corrected = file_path.split('/')[-1].split('.')[0]
        local_dir = f'{dir1_url}/{year_str}/{dir2_url}/{fecha_str_dir}/{nombre_archivo_corrected}'
        
        df_demandas.to_csv(f"{local_dir}_DEMANDAS.csv", index=False, sep=getListSeparator())
        print(f'Archivo "{local_dir}_DEMANDAS.csv" guardado exitosamente.')
        print("----------------------------------\n")

        # Concatenar el dataframe de demandas del dia con el resto
        df_calculos = pd.concat([df_calculos, df_demandas], ignore_index=True)

        fecha_actual += timedelta(days=1)
        
    # Save the final result to a CSV file
    df_calculos.to_csv(f"RESUMEN_DEMANDAS.csv", index=False, sep=getListSeparator())
