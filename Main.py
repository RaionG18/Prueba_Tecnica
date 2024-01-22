from funciones import *

# Configuración directorio WEB
base_url = 'https://www.amm.org.gt/pdfs2/programas_despacho/'
dir1_url = '01_PROGRAMAS_DE_DESPACHO_DIARIO'
dir2_url = '01_PROGRAMAS_DE_DESPACHO_DIARIO'

# Configura la localización en español
locale.setlocale(locale.LC_TIME, 'es_ES.utf8')


if __name__ == "__main__":
    # Define las fechas iniciales y finales
    fecha_inicial = datetime(2023, 1, 1)
    fecha_final = datetime(2023, 6, 30)

    # Descargar archivos para cada día en el rango especificado
    fecha_actual = fecha_inicial

    df_calculos = pd.DataFrame()
    while fecha_actual <= fecha_final:
        # Formatea la fecha actual
        fecha_str_doc = fecha_actual.strftime("%d%m%Y")
        fecha_str_dir = fecha_actual.strftime("%m_%B").upper()
        year_str = fecha_actual.strftime("%Y")

        url_dir = f'{dir1_url}/{year_str}/{dir2_url}/{fecha_str_dir}/WEB{fecha_str_doc}.xlsx'

        # Construye la URL
        url_descarga = urljoin(base_url, url_dir)

        # Construye el directorio
        os.makedirs(os.path.dirname(url_dir), exist_ok=True)
        nombre_local_excel = url_dir

        # Descargar el archivo descomentando la siguiente línea cuando estés listo
        df_demandas = descargar_excel(url_descarga, nombre_local_excel, fecha_actual.date())

        # Agrega un retraso entre solicitudes para simular comportamiento humano
        # time.sleep(1)  # Ajusta el valor según sea necesario

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

        df_demandas['Indicador'] = (df_demandas['POE'] > df_demandas['Costo']).astype(int)
        df_demandas['Liquidacion POE'] = df_demandas['POE']*df_demandas['generacion']
        df_demandas['Liquidacion CVG'] = df_demandas['Costo']*df_demandas['generacion']
        df_demandas['Agente A'] = df_demandas['Indicador']*(df_demandas['Liquidacion POE'] - df_demandas['Liquidacion CVG'])
        df_demandas['Agente B'] = (df_demandas['Indicador']-1)*(df_demandas['Liquidacion POE'] - df_demandas['Liquidacion CVG'])

        nombre_archivo_corrected = nombre_local_excel.replace(".xlsx", "")
        df_demandas.to_csv(f"{nombre_archivo_corrected}_DEMANDAS.csv", index=False, sep=getListSeparator())
        print(f'Archivo "{nombre_archivo_corrected}_DEMANDAS.csv" guardado exitosamente.')
        print("----------------------------------")

        df_calculos = pd.concat([df_calculos, df_demandas], ignore_index=True)
        fecha_actual += timedelta(days=1)

    df_calculos.to_csv(f"RESUMEN_DEMANDAS.csv", index=False, sep=getListSeparator())
