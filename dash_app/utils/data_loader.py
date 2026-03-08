"""
data_loader.py - Carga y preprocesamiento de los CSV de ocupación académica.
"""
import pandas as pd
import os

def cargar_datos(base_path='BBDD'):
    """Carga los 3 CSV necesarios y aplica preprocesamiento."""
    
    prog_path = os.path.join(base_path, 'PROG_DETALLADA_PARA_HEATMAP.csv')
    bloques_path = os.path.join(base_path, 'BLOQUES_ESTANDAR_01122025.csv')
    capacidad_path = os.path.join(base_path, 'PLANTA_FISICA_PREPROCESADA.csv')
    
    # 1. Programación detallada
    df_aperturado = pd.read_csv(prog_path, sep=';', encoding='latin-1')
    
    # 2. Bloques estándar
    df2 = pd.read_csv(bloques_path, sep=';', encoding='latin-1')
    df2.rename(columns={df2.columns[0]: 'BLOQUE'}, inplace=True)
    df2['HORA INICIO'] = pd.to_datetime(df2['HORA INICIO'], format='%H:%M:%S', errors='coerce').dt.strftime('%H:%M:%S')
    df2['HORA FIN'] = pd.to_datetime(df2['HORA FIN'], format='%H:%M:%S', errors='coerce').dt.strftime('%H:%M:%S')
    
    # 3. Capacidad de salas (planta física)
    df3 = pd.read_csv(capacidad_path, sep=';', encoding='utf-8')
    
    return df_aperturado, df2, df3


def obtener_opciones_filtros(df_aperturado, df3):
    """Genera las listas de opciones para cada filtro."""
    
    campus_edificios = sorted(df3['Campus Edificio'].dropna().unique(), key=str)
    clasificaciones_capacidad = sorted(df3['CLASIFICACIÓN'].dropna().unique(), key=str)
    clasificaciones_ocupacion = sorted(df_aperturado['CLASIFICACIÓN'].dropna().unique(), key=str)
    anios = sorted(df_aperturado['AÑO'].dropna().unique(), key=str)
    campus_list = sorted(df_aperturado['CAMPUS'].dropna().unique(), key=str)
    periodos = sorted(df_aperturado['PERIODO'].dropna().unique(), key=str)
    desc_carreras = sorted(df_aperturado['DESCRIPCIÓN CARRERA_MATERIA'].dropna().unique(), key=str)
    
    # Bloques horarios
    bloques_inicio = sorted(df_aperturado['HR_INICIO_AUX_STR'].dropna().astype(str).unique(), key=str)
    bloques_fin = sorted(df_aperturado['HR_FIN_AUX_STR'].dropna().astype(str).unique(), key=str)
    
    # Días
    orden_dias = ['LUNES', 'MARTES', 'MIERCOLES', 'JUEVES', 'VIERNES', 'SABADO', 'DOMINGO']
    dias_disponibles = [d for d in orden_dias if d in df_aperturado['DIA_SESION'].unique()]
    
    # Fechas
    fechas_ini = pd.to_datetime(df_aperturado['FECHA_INI2'], origin='1899-12-30', unit='D', errors='coerce')
    fechas_fin = pd.to_datetime(df_aperturado['FECHA_TERM2'], origin='1899-12-30', unit='D', errors='coerce')
    fechas_ini_v = fechas_ini[fechas_ini.notna() & (fechas_ini.dt.year >= 1900)]
    fechas_fin_v = fechas_fin[fechas_fin.notna() & (fechas_fin.dt.year >= 1900)]
    
    fecha_min = fechas_ini_v.min().date() if not fechas_ini_v.empty else pd.Timestamp('2020-01-01').date()
    fecha_max = fechas_fin_v.max().date() if not fechas_fin_v.empty else pd.Timestamp('2030-12-31').date()
    
    return {
        'campus_edificios': campus_edificios,
        'clasificaciones_capacidad': clasificaciones_capacidad,
        'clasificaciones_ocupacion': clasificaciones_ocupacion,
        'anios': anios,
        'campus_list': campus_list,
        'periodos': periodos,
        'desc_carreras': desc_carreras,
        'bloques_inicio': bloques_inicio,
        'bloques_fin': bloques_fin,
        'dias_disponibles': dias_disponibles,
        'fecha_min': fecha_min,
        'fecha_max': fecha_max,
    }
