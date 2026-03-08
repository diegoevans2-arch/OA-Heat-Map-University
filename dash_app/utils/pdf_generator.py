"""
pdf_generator.py - Genera el PDF de Resumen Ejecutivo.
Adaptado del notebook para funcionar de forma independiente en la app Dash.
"""
import pandas as pd
import os
import tempfile
from pathlib import Path
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch, cm
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table,
                                 TableStyle, PageBreak, HRFlowable, KeepTogether, Image)
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT, TA_JUSTIFY


def get_download_folder():
    """Detecta la carpeta de descargas del usuario."""
    carpeta = str(Path.home() / 'Downloads')
    if not os.path.exists(carpeta):
        carpeta = str(Path.home() / 'Descargas')
    if not os.path.exists(carpeta):
        carpeta = str(Path.home())
    return carpeta
def generar_resumen_ejecutivo_pdf(df_filtrado, df3_ref, df2_ref, filtros_dict, tipo_grafico="Heatmap", fig_imagen=None):
    """
    Genera un PDF con formato de Resumen Ejecutivo de ocupación.
    
    Parámetros:
    - df_filtrado: DataFrame con los datos filtrados de ocupación
    - df3_ref: DataFrame de capacidad de salas (df3)
    - df2_ref: DataFrame de bloques estándar (df2)
    - filtros_dict: dict con los filtros aplicados
    - tipo_grafico: str - "Heatmap" o "Sunburst"
    - fig_imagen: plotly Figure - gráfico a insertar como imagen en el PDF
    """
    
    if df_filtrado is None or df_filtrado.empty:
        print("⚠️ No hay datos filtrados. Primero genera el gráfico.")
        return
    
    timestamp = pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')
    download_folder = get_download_folder()
    filename = os.path.join(download_folder, f'Resumen_Ejecutivo_{tipo_grafico}_{timestamp}.pdf')
    
    # =============================================
    # PREPARACIÓN DE DATOS
    # =============================================
    df_work = df_filtrado.copy()
    for col in ['CAMPUS_AUX', 'EDIFICIO', 'SALA', 'CLASIFICACIÓN', 'CAMPUS']:
        if col in df_work.columns:
            df_work[col] = df_work[col].astype(str)
    
    # Filtrar solo registros con SALA válida (no vacía, no null, no 'nan')
    df_work = df_work[
        (df_work['SALA'].notna()) & 
        (df_work['SALA'] != '') & 
        (df_work['SALA'] != 'nan') &
        (df_work['SALA'].str.strip() != '')
    ].copy()
    
    if df_work.empty:
        print("⚠️ No hay registros con sala válida.")
        return
    
    col_ocup = 'OCUPACIÓN_SALA' if 'OCUPACIÓN_SALA' in df_work.columns else None
    if not col_ocup:
        print("⚠️ No se encontró columna OCUPACIÓN_SALA.")
        return
    
    # --- Calcular total_capacidad del campus (desde df3 con mismos filtros) ---
    campus_edificio_list = filtros_dict.get('campus_edificio_list', [])
    excl_cap = filtros_dict.get('clasificacion_excluir_capacidad', [])
    
    filtro_df3 = df3_ref['Campus Edificio'].isin(campus_edificio_list)
    if excl_cap:
        filtro_df3 = filtro_df3 & (~df3_ref['CLASIFICACIÓN'].isin(excl_cap))
    total_capacidad = df3_ref.loc[filtro_df3, 'Capacidad Máxima Sala'].sum()
    if total_capacidad == 0:
        total_capacidad = 1  # evitar division por cero
    
    # --- Tabla de % ocupación por sala (campo OCUPACIÓN_SALA promedio) ---
    group_cols = ['CAMPUS_AUX', 'EDIFICIO', 'CLASIFICACIÓN', 'SALA']
    if 'Rango_Capacidad' in df_work.columns:
        group_cols = ['CAMPUS_AUX', 'EDIFICIO', 'CLASIFICACIÓN', 'Rango_Capacidad', 'SALA']
    
    tabla_sala = df_work.groupby(group_cols).agg(
        Ocupacion_Prom=(col_ocup, 'mean'),
        N_Registros=(col_ocup, 'count')
    ).reset_index()
    tabla_sala['Ocupacion_Pct'] = (tabla_sala['Ocupacion_Prom'] * 100).round(2)
    
    # Promedio por campus = promedio simple de OCUPACIÓN_SALA de todos los registros con sala válida
    promedio_campus = df_work.groupby('CAMPUS_AUX')[col_ocup].mean().to_dict()
    promedio_campus = {k: round(v * 100, 2) for k, v in promedio_campus.items()}
    
    # Ordenar
    tabla_sala = tabla_sala.sort_values('Ocupacion_Pct', ascending=False).reset_index(drop=True)
    
    # =============================================
    # ESTILOS PDF
    # =============================================
    doc = SimpleDocTemplate(filename, pagesize=letter,
                            topMargin=1.5*cm, bottomMargin=1.5*cm,
                            leftMargin=1.5*cm, rightMargin=1.5*cm)
    
    styles = getSampleStyleSheet()
    
    style_titulo = ParagraphStyle('TituloEjecutivo', parent=styles['Title'],
                                   fontSize=18, textColor=colors.HexColor('#1a237e'),
                                   spaceAfter=6, alignment=TA_CENTER)
    
    style_subtitulo = ParagraphStyle('Subtitulo', parent=styles['Heading2'],
                                      fontSize=13, textColor=colors.HexColor('#283593'),
                                      spaceBefore=14, spaceAfter=6)
    
    style_seccion = ParagraphStyle('Seccion', parent=styles['Heading3'],
                                    fontSize=11, textColor=colors.HexColor('#1565c0'),
                                    spaceBefore=10, spaceAfter=4)
    
    style_normal = ParagraphStyle('NormalCustom', parent=styles['Normal'],
                                   fontSize=9, leading=12, alignment=TA_JUSTIFY)
    
    style_small = ParagraphStyle('Small', parent=styles['Normal'],
                                  fontSize=8, leading=10, textColor=colors.HexColor('#555555'))
    
    style_alerta = ParagraphStyle('Alerta', parent=styles['Normal'],
                                   fontSize=9, leading=12, textColor=colors.HexColor('#b71c1c'),
                                   fontName='Helvetica-Bold')
    
    style_parrafo_resumen = ParagraphStyle('ParrafoResumen', parent=styles['Normal'],
                                            fontSize=9, leading=13, alignment=TA_JUSTIFY,
                                            spaceBefore=4, spaceAfter=4)
    
    story = []
    
    # =============================================
    # ENCABEZADO
    # =============================================
    story.append(Paragraph("RESUMEN EJECUTIVO", style_titulo))
    story.append(Paragraph("Análisis de Ocupación de Espacios Académicos",
                           ParagraphStyle('Sub', parent=styles['Normal'],
                                          fontSize=12, alignment=TA_CENTER,
                                          textColor=colors.HexColor('#455a64'), spaceAfter=4)))
    
    fecha_gen = pd.Timestamp.now().strftime('%d/%m/%Y %H:%M')
    story.append(Paragraph(f"Generado: {fecha_gen} | Fuente: {tipo_grafico}",
                           ParagraphStyle('Fecha', parent=styles['Normal'],
                                          fontSize=8, alignment=TA_CENTER,
                                          textColor=colors.HexColor('#888888'), spaceAfter=8)))
    
    story.append(HRFlowable(width="100%", thickness=2, color=colors.HexColor('#1a237e')))
    story.append(Spacer(1, 10))
    
    # =============================================
    # 1. RESUMEN EJECUTIVO EN PÁRRAFO
    # =============================================
    story.append(Paragraph("1. Resumen Ejecutivo", style_subtitulo))
    
    # Construir párrafo dinámico - Solo registros donde CAMPUS == CAMPUS_AUX (campus de análisis)
    campus_list_filtro = filtros_dict.get('campus_list', [])
    anio = filtros_dict.get('anio', '-')
    
    # Para el resumen, filtrar solo registros del campus de análisis
    df_resumen = df_work[df_work['CAMPUS'] == df_work['CAMPUS_AUX']].copy() if 'CAMPUS' in df_work.columns else df_work.copy()
    
    # Datos para el resumen (solo campus de análisis: CAMPUS == CAMPUS_AUX)
    tabla_sala_resumen = tabla_sala.copy()
    if 'Rango_Capacidad' in tabla_sala_resumen.columns:
        tabla_sala_resumen = tabla_sala_resumen[tabla_sala_resumen['Rango_Capacidad'].astype(str) != 'Sin Capacidad']
    n_salas_total = tabla_sala_resumen['SALA'].nunique()
    n_criticas = len(tabla_sala_resumen[tabla_sala_resumen['Ocupacion_Pct'] <= 20])
    
    # Bloques con menor ocupación (excluyendo sab/dom, solo campus de análisis)
    df_bloques_resumen = df_resumen[~df_resumen['DIA_SESION'].isin(['SABADO', 'DOMINGO'])].copy() if 'DIA_SESION' in df_resumen.columns else df_resumen.copy()
    
    bloques_bajos_resumen = []
    if 'HR_INICIO_AUX_STR' in df_bloques_resumen.columns:
        # Filtrar solo bloques de pregrado (08:00:00 a 18:55:00, bloques 1 a 14)
        df_bloques_resumen = df_bloques_resumen[
            (df_bloques_resumen['HR_INICIO_AUX_STR'].astype(str) >= '08:00:00') &
            (df_bloques_resumen['HR_INICIO_AUX_STR'].astype(str) <= '18:55:00')
        ].copy()
        
        bloq_r = df_bloques_resumen.groupby('HR_INICIO_AUX_STR')['INSCRITOS'].sum().reset_index()
        # contar días L-V por bloque
        dias_por_bloque_r = df_bloques_resumen.groupby('HR_INICIO_AUX_STR')['DIA_SESION'].nunique().reset_index()
        dias_por_bloque_r.columns = ['HR_INICIO_AUX_STR', 'N_DIAS']
        bloq_r = bloq_r.merge(dias_por_bloque_r, on='HR_INICIO_AUX_STR', how='left')
        bloq_r['N_DIAS'] = bloq_r['N_DIAS'].fillna(1).astype(int)
        bloq_r['Ocup_Pct'] = ((bloq_r['INSCRITOS'] / (total_capacidad * bloq_r['N_DIAS'])) * 100).round(2)
        bloq_r = bloq_r.sort_values('Ocup_Pct')
        bloques_bajos_resumen = bloq_r[bloq_r['Ocup_Pct'] < 20].head(5)['HR_INICIO_AUX_STR'].tolist()
    
    # Salas bajo promedio (excluyendo Sin Capacidad)
    n_bajo_prom_total = 0
    for campus_val in promedio_campus:
        prom_c = promedio_campus[campus_val]
        sub = tabla_sala_resumen[tabla_sala_resumen['CAMPUS_AUX'] == campus_val]
        n_bajo_prom_total += len(sub[sub['Ocupacion_Pct'] < prom_c])
    
    # Construir texto
    parrafos_resumen = []
    
    # Solo campus donde CAMPUS == CAMPUS_AUX (campus de análisis)
    campus_analisis = sorted(set(df_resumen['CAMPUS_AUX'].unique()) & set(promedio_campus.keys()))
    
    for campus_val in campus_analisis:
        prom_c = promedio_campus[campus_val]
        sub_campus = tabla_sala_resumen[tabla_sala_resumen['CAMPUS_AUX'] == campus_val]
        n_salas_campus = len(sub_campus)
        n_criticas_campus = len(sub_campus[sub_campus['Ocupacion_Pct'] <= 20])
        n_bajo_campus = len(sub_campus[sub_campus['Ocupacion_Pct'] < prom_c])
        
        texto = (f"El campus <b>{campus_val}</b> presenta una ocupación promedio de <b>{prom_c:.2f}%</b> "
                 f"sobre un total de {n_salas_campus} salas analizadas. ")
        
        if prom_c < 30:
            texto += "Este nivel indica una subutilización significativa de la infraestructura disponible. "
        elif prom_c < 60:
            texto += "Existen oportunidades de mejora en la asignación de espacios. "
        else:
            texto += "El nivel de uso se encuentra en un rango aceptable. "
        
        if n_criticas_campus > 0:
            texto += (f"Se identifican <b>{n_criticas_campus} salas en situación crítica</b> (ocupación igual o "
                      f"menor al 20%), las cuales requieren atención inmediata para evaluar su reasignación "
                      f"o reconversión. ")
        
        if n_bajo_campus > 0:
            texto += (f"Adicionalmente, {n_bajo_campus} salas se encuentran bajo el promedio del campus. ")
        
        parrafos_resumen.append(texto)
    
    # Info de bloques bajos
    if bloques_bajos_resumen:
        bloques_str = ', '.join(bloques_bajos_resumen)
        parrafos_resumen.append(
            f"En cuanto a la distribución horaria de <b>pregrado</b> (bloques 1 a 14, de 08:00 a 18:55), "
            f"se identifican los bloques <b>{bloques_str}</b> con baja frecuencia de uso (bajo el 20% de "
            f"ocupación), lo que representa una oportunidad para optimizar la programación académica "
            f"o destinar esos horarios a actividades complementarias."
        )
    
    # Párrafo de flujo por bloque (pregrado 1-14) y por día
    # Flujo por bloque pregrado: SUM(INSCRITOS) / total_capacidad
    if 'HR_INICIO_AUX_STR' in df_bloques_resumen.columns and not df_bloques_resumen.empty:
        flujo_bloque_r = df_bloques_resumen.groupby('HR_INICIO_AUX_STR')['INSCRITOS'].sum().reset_index()
        flujo_bloque_r['Flujo'] = (flujo_bloque_r['INSCRITOS'] / total_capacidad).round(2)
        
        if not flujo_bloque_r.empty:
            bloque_max = flujo_bloque_r.loc[flujo_bloque_r['Flujo'].idxmax()]
            bloque_min = flujo_bloque_r.loc[flujo_bloque_r['Flujo'].idxmin()]
            
            parrafos_resumen.append(
                f"En términos de flujo de estudiantes en bloques de pregrado, el bloque con "
                f"<b>mayor flujo</b> es el de las <b>{bloque_max['HR_INICIO_AUX_STR']}</b> "
                f"({bloque_max['Flujo']:.2f} veces la capacidad del campus), mientras que el de "
                f"<b>menor flujo</b> corresponde a las <b>{bloque_min['HR_INICIO_AUX_STR']}</b> "
                f"({bloque_min['Flujo']:.2f} veces la capacidad)."
            )
    
    # Flujo por día (L-V, sin filtro de bloques pregrado, desde df_resumen completo)
    df_dias_resumen = df_resumen[~df_resumen['DIA_SESION'].isin(['SABADO', 'DOMINGO'])].copy() if 'DIA_SESION' in df_resumen.columns else pd.DataFrame()
    
    if not df_dias_resumen.empty and 'DIA_SESION' in df_dias_resumen.columns:
        flujo_dia_r = df_dias_resumen.groupby('DIA_SESION')['INSCRITOS'].sum().reset_index()
        flujo_dia_r['Flujo'] = (flujo_dia_r['INSCRITOS'] / total_capacidad).round(2)
        
        if not flujo_dia_r.empty:
            dia_max = flujo_dia_r.loc[flujo_dia_r['Flujo'].idxmax()]
            dia_min = flujo_dia_r.loc[flujo_dia_r['Flujo'].idxmin()]
            
            parrafos_resumen.append(
                f"Respecto al flujo diario, el día con <b>mayor flujo</b> de estudiantes es "
                f"<b>{dia_max['DIA_SESION']}</b> ({dia_max['Flujo']:.2f} veces la capacidad), "
                f"y el de <b>menor flujo</b> es <b>{dia_min['DIA_SESION']}</b> "
                f"({dia_min['Flujo']:.2f} veces la capacidad)."
            )
    
    # Párrafo de rectificación de datos
    # Contar salas sin capacidad y sin clasificación
    n_sin_cap_resumen = 0
    n_sin_clasif_resumen = 0
    if 'Rango_Capacidad' in df_work.columns:
        n_sin_cap_resumen = df_work[df_work['Rango_Capacidad'].astype(str) == 'Sin Capacidad']['SALA'].nunique()
    if 'CLASIFICACIÓN' in df_work.columns:
        clasif_invalidas = ['Sin Clasificación', 'Sin Clasificacion', 'nan', '']
        n_sin_clasif_resumen = df_work[df_work['CLASIFICACIÓN'].astype(str).str.strip().isin(clasif_invalidas)]['SALA'].nunique()
    
    if n_sin_cap_resumen > 0 or n_sin_clasif_resumen > 0:
        texto_rect = "Respecto a la calidad de los datos, se detectó que "
        partes_rect = []
        if n_sin_cap_resumen > 0:
            partes_rect.append(f"<b>{n_sin_cap_resumen} salas</b> no cuentan con capacidad registrada o tienen capacidad 0")
        if n_sin_clasif_resumen > 0:
            partes_rect.append(f"<b>{n_sin_clasif_resumen} salas</b> figuran con 'Sin Clasificación'")
        texto_rect += ' y '.join(partes_rect) + ". "
        texto_rect += ("Esta situación afecta directamente la precisión de los indicadores de ocupación "
                       "presentados en este informe. Se recomienda que el área responsable de planta física "
                       "realice la rectificación de estos registros a la brevedad.")
        parrafos_resumen.append(texto_rect)
    
    for p in parrafos_resumen:
        story.append(Paragraph(p, style_parrafo_resumen))
    
    story.append(Spacer(1, 6))
    
    # =============================================
    # INSERTAR IMAGEN DEL GRÁFICO
    # =============================================
    if fig_imagen is not None:
        try:
            # Exportar gráfico a PNG temporal
            img_temp_path = os.path.join(tempfile.gettempdir(), f'grafico_{tipo_grafico}_{pd.Timestamp.now().strftime("%Y%m%d_%H%M%S")}.png')
            fig_imagen.write_image(img_temp_path, format='png', width=1000, height=700, scale=2)
            
            # Calcular tamaño para que quepa en la página (ancho máximo disponible)
            page_w = letter[0] - 3*cm  # ancho disponible
            img_w = page_w
            img_h = img_w * 0.7  # proporción 10:7
            
            # Ajustar si es muy alto para la primera página
            max_h = 12*cm
            if img_h > max_h:
                img_h = max_h
                img_w = img_h / 0.7
            
            story.append(Paragraph(f"<b>Gráfico: {tipo_grafico}</b>",
                         ParagraphStyle('GrafTitulo', parent=styles['Normal'],
                                        fontSize=10, alignment=TA_CENTER, spaceBefore=2, spaceAfter=4)))
            story.append(Image(img_temp_path, width=img_w, height=img_h))
            story.append(Spacer(1, 8))
            
            print(f"   📊 Imagen del gráfico insertada en el PDF")
        except Exception as e:
            print(f"   ⚠️ No se pudo insertar la imagen del gráfico: {str(e)}")
            print(f"      Verifique que 'kaleido' esté instalado (pip install kaleido)")
    
    story.append(Spacer(1, 6))
    
    # =============================================
    # 2. FILTROS USADOS
    # =============================================
    story.append(Paragraph("2. Filtros Aplicados", style_subtitulo))
    
    filtros_data = [['Parámetro', 'Valor']]
    filtros_data.append(['AÑO', str(filtros_dict.get('anio', '-'))])
    filtros_data.append(['Campus', ', '.join([str(c) for c in filtros_dict.get('campus_list', [])])])
    filtros_data.append(['Edificios', ', '.join([str(e) for e in filtros_dict.get('campus_edificio_list', [])])])
    
    periodos = filtros_dict.get('periodo_list', [])
    filtros_data.append(['Período Académico', ', '.join([str(p) for p in periodos]) if periodos else 'Todos'])
    
    carreras = filtros_dict.get('desc_carrera_materia_list', [])
    carr_txt = ', '.join([str(c) for c in carreras[:5]]) + ('...' if len(carreras) > 5 else '') if carreras else 'Todas'
    filtros_data.append(['Carrera/Materia', carr_txt])
    
    fecha_ini = filtros_dict.get('fecha_inicio')
    fecha_f = filtros_dict.get('fecha_fin')
    filtros_data.append(['Rango Fechas', f'{fecha_ini} a {fecha_f}' if fecha_ini and fecha_f else 'Sin filtro'])
    
    excl_cap_f = filtros_dict.get('clasificacion_excluir_capacidad', [])
    filtros_data.append(['Clasif. Excluidas (Cap.)', ', '.join([str(c) for c in excl_cap_f]) if excl_cap_f else 'Ninguna'])
    
    excl_ocu_f = filtros_dict.get('clasificacion_excluir_ocupacion', [])
    filtros_data.append(['Clasif. Excluidas (Prog.)', ', '.join([str(c) for c in excl_ocu_f]) if excl_ocu_f else 'Ninguna'])
    
    if 'dias_sesion_list' in filtros_dict and filtros_dict['dias_sesion_list']:
        filtros_data.append(['Dias Sesion', ', '.join([str(d) for d in filtros_dict['dias_sesion_list']])])
    if 'bloque_inicio' in filtros_dict and filtros_dict['bloque_inicio']:
        filtros_data.append(['Bloque Horario', f"{filtros_dict['bloque_inicio']} - {filtros_dict.get('bloque_fin', '')}"])
    
    filtros_data.append(['Capacidad Total Campus', f"{total_capacidad:,.0f}"])
    
    t_filtros = Table(filtros_data, colWidths=[150, 340])
    t_filtros.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1a237e')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 8),
        ('BACKGROUND', (0, 1), (0, -1), colors.HexColor('#e8eaf6')),
        ('FONTNAME', (0, 1), (0, -1), 'Helvetica-Bold'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#bdbdbd')),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('LEFTPADDING', (0, 0), (-1, -1), 6),
        ('RIGHTPADDING', (0, 0), (-1, -1), 6),
        ('TOPPADDING', (0, 0), (-1, -1), 3),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
    ]))
    story.append(t_filtros)
    story.append(Spacer(1, 12))
    
    # =============================================
    # 3. TABLA DE % OCUPACIÓN - TODAS LAS SALAS
    # =============================================
    story.append(Paragraph("3. Porcentaje de Ocupación por Campus, Edificio, Clasificación, Rango y Sala", style_subtitulo))
    
    has_rango = 'Rango_Capacidad' in tabla_sala.columns
    
    # Encabezado
    if has_rango:
        tabla_data = [['Campus', 'Edificio', 'Clasificación', 'Rango Cap.', 'Sala', '% Ocup.']]
    else:
        tabla_data = [['Campus', 'Edificio', 'Clasificación', 'Sala', '% Ocup.']]
    
    for _, row in tabla_sala.iterrows():
        if has_rango:
            tabla_data.append([
                str(row['CAMPUS_AUX']), str(row['EDIFICIO']), str(row['CLASIFICACIÓN']),
                str(row.get('Rango_Capacidad', '')), str(row['SALA']),
                f"{row['Ocupacion_Pct']:.2f}%"
            ])
        else:
            tabla_data.append([
                str(row['CAMPUS_AUX']), str(row['EDIFICIO']), str(row['CLASIFICACIÓN']),
                str(row['SALA']), f"{row['Ocupacion_Pct']:.2f}%"
            ])
    
    if has_rango:
        col_widths_tabla = [45, 50, 70, 65, 50, 50]
    else:
        col_widths_tabla = [50, 60, 80, 55, 55]
    
    t_ocup = Table(tabla_data, colWidths=col_widths_tabla, repeatRows=1)
    
    estilo_tabla = [
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1a237e')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 6.5),
        ('GRID', (0, 0), (-1, -1), 0.3, colors.HexColor('#cccccc')),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('ALIGN', (-1, 0), (-1, -1), 'CENTER'),
        ('LEFTPADDING', (0, 0), (-1, -1), 3),
        ('RIGHTPADDING', (0, 0), (-1, -1), 3),
        ('TOPPADDING', (0, 0), (-1, -1), 2),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
    ]
    
    # Colores: ROJO <= 20%, AMARILLO < 60%, VERDE >= 60%
    for i in range(1, len(tabla_data)):
        try:
            val = float(tabla_data[i][-1].replace('%', '').replace(',', '.'))
            if val <= 20:
                estilo_tabla.append(('BACKGROUND', (0, i), (-1, i), colors.HexColor('#ffcdd2')))
            elif val < 60:
                estilo_tabla.append(('BACKGROUND', (0, i), (-1, i), colors.HexColor('#fff9c4')))
            else:
                estilo_tabla.append(('BACKGROUND', (0, i), (-1, i), colors.HexColor('#e8f5e9')))
        except:
            pass
    
    t_ocup.setStyle(TableStyle(estilo_tabla))
    story.append(t_ocup)
    story.append(Spacer(1, 4))
    
    # Leyenda
    leyenda = Table([['Rojo: <= 20%', 'Amarillo: 20% - 59%', 'Verde: >= 60%']], colWidths=[130, 130, 130])
    leyenda.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (0, 0), colors.HexColor('#ffcdd2')),
        ('BACKGROUND', (1, 0), (1, 0), colors.HexColor('#fff9c4')),
        ('BACKGROUND', (2, 0), (2, 0), colors.HexColor('#e8f5e9')),
        ('FONTSIZE', (0, 0), (-1, -1), 7),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('BOX', (0, 0), (-1, -1), 0.5, colors.grey),
    ]))
    story.append(leyenda)
    story.append(Spacer(1, 12))
    
    # =============================================
    # 4. SALAS BAJO EL PROMEDIO DE SU CAMPUS
    # =============================================
    story.append(Paragraph("4. Salas Bajo el Promedio de su Campus", style_subtitulo))
    
    salas_bajo_prom = []
    for _, row in tabla_sala.iterrows():
        campus = row['CAMPUS_AUX']
        prom_campus_val = promedio_campus.get(campus, 0)
        if row['Ocupacion_Pct'] < prom_campus_val:
            entry = {
                'Campus': campus, 'Edificio': row['EDIFICIO'],
                'Clasificación': row['CLASIFICACIÓN'], 'Sala': row['SALA'],
                'Ocupación': row['Ocupacion_Pct'], 'Promedio_Campus': prom_campus_val
            }
            if has_rango:
                entry['Rango_Capacidad'] = row.get('Rango_Capacidad', '')
            salas_bajo_prom.append(entry)
    
    df_bajo_prom = pd.DataFrame(salas_bajo_prom)
    
    # Excluir salas con Rango_Capacidad = "Sin Capacidad"
    if not df_bajo_prom.empty and 'Rango_Capacidad' in df_bajo_prom.columns:
        df_bajo_prom = df_bajo_prom[df_bajo_prom['Rango_Capacidad'].astype(str) != 'Sin Capacidad']
    
    if not df_bajo_prom.empty:
        for campus_val in sorted(df_bajo_prom['Campus'].unique()):
            sub = df_bajo_prom[df_bajo_prom['Campus'] == campus_val]
            prom_c = promedio_campus.get(campus_val, 0)
            story.append(Paragraph(
                f"Campus {campus_val} (Promedio: {prom_c:.2f}%) - {len(sub)} salas bajo promedio",
                style_seccion))
            
            if has_rango:
                bajo_data = [['Edificio', 'Clasificación', 'Rango Cap.', 'Sala', '% Ocup.', 'Diferencia']]
            else:
                bajo_data = [['Edificio', 'Clasificación', 'Sala', '% Ocup.', 'Diferencia']]
            
            for _, r in sub.sort_values('Ocupación').iterrows():
                diff = r['Ocupación'] - r['Promedio_Campus']
                if has_rango:
                    bajo_data.append([
                        str(r['Edificio']), str(r['Clasificación']), str(r.get('Rango_Capacidad', '')),
                        str(r['Sala']), f"{r['Ocupación']:.2f}%", f"{diff:.2f}%"
                    ])
                else:
                    bajo_data.append([
                        str(r['Edificio']), str(r['Clasificación']),
                        str(r['Sala']), f"{r['Ocupación']:.2f}%", f"{diff:.2f}%"
                    ])
            
            if has_rango:
                cw_bajo = [55, 70, 60, 50, 50, 50]
            else:
                cw_bajo = [70, 90, 60, 60, 65]
            
            t_bajo = Table(bajo_data, colWidths=cw_bajo, repeatRows=1)
            t_bajo.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#e65100')),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, -1), 6.5),
                ('GRID', (0, 0), (-1, -1), 0.3, colors.HexColor('#cccccc')),
                ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor('#fff3e0')),
                ('ALIGN', (-2, 0), (-1, -1), 'CENTER'),
                ('TOPPADDING', (0, 0), (-1, -1), 2),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
                ('LEFTPADDING', (0, 0), (-1, -1), 3),
                ('RIGHTPADDING', (0, 0), (-1, -1), 3),
            ]))
            story.append(t_bajo)
            story.append(Spacer(1, 6))
    else:
        story.append(Paragraph("No se encontraron salas bajo el promedio de su campus.", style_normal))
    
    story.append(Spacer(1, 8))
    
    # =============================================
    # 5. SALAS CRÍTICAS (OCUPACIÓN <= 20%)
    # =============================================
    story.append(Paragraph("5. Salas Críticas (Ocupación Igual o Menor al 20%)", style_subtitulo))
    
    salas_criticas = tabla_sala[tabla_sala['Ocupacion_Pct'] <= 20].copy()
    
    # Excluir salas con Rango_Capacidad = "Sin Capacidad"
    if 'Rango_Capacidad' in salas_criticas.columns:
        salas_criticas = salas_criticas[salas_criticas['Rango_Capacidad'].astype(str) != 'Sin Capacidad']
    
    if not salas_criticas.empty:
        story.append(Paragraph(
            f"Se identificaron {len(salas_criticas)} salas con ocupación crítica (igual o menor al 20%).",
            style_alerta))
        story.append(Spacer(1, 4))
        
        if has_rango:
            crit_data = [['Campus', 'Edificio', 'Clasificación', 'Rango Cap.', 'Sala', '% Ocup.']]
        else:
            crit_data = [['Campus', 'Edificio', 'Clasificación', 'Sala', '% Ocup.']]
        
        for _, r in salas_criticas.sort_values('Ocupacion_Pct').iterrows():
            if has_rango:
                crit_data.append([
                    str(r['CAMPUS_AUX']), str(r['EDIFICIO']), str(r['CLASIFICACIÓN']),
                    str(r.get('Rango_Capacidad', '')), str(r['SALA']), f"{r['Ocupacion_Pct']:.2f}%"
                ])
            else:
                crit_data.append([
                    str(r['CAMPUS_AUX']), str(r['EDIFICIO']), str(r['CLASIFICACIÓN']),
                    str(r['SALA']), f"{r['Ocupacion_Pct']:.2f}%"
                ])
        
        if has_rango:
            cw_crit = [45, 50, 70, 65, 50, 50]
        else:
            cw_crit = [50, 60, 80, 55, 55]
        
        t_crit = Table(crit_data, colWidths=cw_crit, repeatRows=1)
        t_crit.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#b71c1c')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 6.5),
            ('GRID', (0, 0), (-1, -1), 0.3, colors.HexColor('#cccccc')),
            ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor('#ffebee')),
            ('ALIGN', (-1, 0), (-1, -1), 'CENTER'),
            ('TOPPADDING', (0, 0), (-1, -1), 2),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
            ('LEFTPADDING', (0, 0), (-1, -1), 3),
            ('RIGHTPADDING', (0, 0), (-1, -1), 3),
        ]))
        story.append(t_crit)
    else:
        story.append(Paragraph("No se encontraron salas con ocupación igual o menor al 20%.", style_normal))
    
    story.append(Spacer(1, 12))
    
    # =============================================
    # 6. BLOQUES HORARIOS DE MENOR OCUPACIÓN
    # =============================================
    story.append(Paragraph("6. Bloques Horarios de Menor Ocupación", style_subtitulo))
    story.append(Paragraph(
        "<i>Nota: Se excluyen los días sábado y domingo del cálculo de bloques. "
        "El % de ocupación se calcula como SUM(Inscritos) / (Capacidad Total Campus x Cantidad de Dias L-V del bloque). "
        "El flujo de estudiantes se calcula como SUM(Inscritos) / Capacidad Total Campus (solo días lunes a viernes).</i>",
        style_small))
    story.append(Spacer(1, 4))
    
    if 'HR_INICIO_AUX_STR' in df_work.columns and 'DIA_SESION' in df_work.columns:
        # Excluir sábado y domingo
        df_bloques = df_work[~df_work['DIA_SESION'].isin(['SABADO', 'DOMINGO'])].copy()
        
        if not df_bloques.empty:
            # Sumatoria de inscritos por bloque
            bloq_inscritos = df_bloques.groupby('HR_INICIO_AUX_STR')['INSCRITOS'].sum().reset_index()
            bloq_inscritos.columns = ['HR_INICIO_AUX_STR', 'SUM_INSCRITOS']
            
            # Contar días únicos L-V por bloque
            dias_por_bloque = df_bloques.groupby('HR_INICIO_AUX_STR')['DIA_SESION'].nunique().reset_index()
            dias_por_bloque.columns = ['HR_INICIO_AUX_STR', 'N_DIAS']
            
            bloq_data_calc = bloq_inscritos.merge(dias_por_bloque, on='HR_INICIO_AUX_STR', how='left')
            bloq_data_calc['N_DIAS'] = bloq_data_calc['N_DIAS'].fillna(1).astype(int)
            
            # % Ocupación = SUM(INSCRITOS) / (total_capacidad * N_DIAS) * 100
            bloq_data_calc['Ocup_Pct'] = ((bloq_data_calc['SUM_INSCRITOS'] / (total_capacidad * bloq_data_calc['N_DIAS'])) * 100).round(2)
            
            # Flujo estudiantes (solo L-V) = SUM(INSCRITOS L-V) / total_capacidad
            df_bloques_lv = df_work[~df_work['DIA_SESION'].isin(['SABADO', 'DOMINGO'])].copy()
            flujo_lv = df_bloques_lv.groupby('HR_INICIO_AUX_STR')['INSCRITOS'].sum().reset_index()
            flujo_lv.columns = ['HR_INICIO_AUX_STR', 'SUM_INSCRITOS_LV']
            flujo_lv['Flujo_Estudiantes'] = (flujo_lv['SUM_INSCRITOS_LV'] / total_capacidad).round(2)
            
            bloq_data_calc = bloq_data_calc.merge(flujo_lv[['HR_INICIO_AUX_STR', 'Flujo_Estudiantes']], 
                                                    on='HR_INICIO_AUX_STR', how='left')
            bloq_data_calc = bloq_data_calc.sort_values('Ocup_Pct').reset_index(drop=True)
            
            bloq_tabla = [['Bloque Horario', '% Ocupación', 'Flujo Estudiantes']]
            for _, r in bloq_data_calc.iterrows():
                bloq_tabla.append([
                    str(r['HR_INICIO_AUX_STR']),
                    f"{r['Ocup_Pct']:.2f}%",
                    f"{r['Flujo_Estudiantes']:.2f}"
                ])
            
            t_bloq = Table(bloq_tabla, colWidths=[120, 110, 110], repeatRows=1)
            t_bloq.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1a237e')),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, -1), 8),
                ('GRID', (0, 0), (-1, -1), 0.4, colors.HexColor('#cccccc')),
                ('ALIGN', (1, 0), (2, -1), 'CENTER'),
                ('TOPPADDING', (0, 0), (-1, -1), 3),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
            ]))
            story.append(t_bloq)
            story.append(Spacer(1, 8))
            
            # Ocupación promedio por día (excluyendo sab y dom, misma lógica)
            story.append(Paragraph("Ocupación promedio por día (Lunes a Viernes):", style_seccion))
            story.append(Paragraph(
                "<i>Nota: Se excluyen sábado y domingo. Flujo de estudiantes calculado solo lunes a viernes.</i>",
                style_small))
            story.append(Spacer(1, 4))
            
            dias_inscritos = df_bloques_lv.groupby('DIA_SESION')['INSCRITOS'].sum().reset_index()
            dias_inscritos.columns = ['DIA_SESION', 'SUM_INSCRITOS']
            
            # Para % ocupación por día: contar bloques únicos por día
            bloques_por_dia = df_bloques_lv.groupby('DIA_SESION')['HR_INICIO_AUX_STR'].nunique().reset_index()
            bloques_por_dia.columns = ['DIA_SESION', 'N_BLOQUES']
            
            dias_calc = dias_inscritos.merge(bloques_por_dia, on='DIA_SESION', how='left')
            dias_calc['N_BLOQUES'] = dias_calc['N_BLOQUES'].fillna(1).astype(int)
            dias_calc['Ocup_Pct'] = ((dias_calc['SUM_INSCRITOS'] / (total_capacidad * dias_calc['N_BLOQUES'])) * 100).round(2)
            dias_calc['Flujo_Estudiantes'] = (dias_calc['SUM_INSCRITOS'] / total_capacidad).round(2)
            
            orden_dias = ['LUNES', 'MARTES', 'MIERCOLES', 'JUEVES', 'VIERNES']
            dias_calc['orden'] = dias_calc['DIA_SESION'].apply(lambda x: orden_dias.index(x) if x in orden_dias else 99)
            dias_calc = dias_calc.sort_values('orden')
            
            dias_tabla = [['Dia', '% Ocupación', 'Flujo Estudiantes']]
            for _, r in dias_calc.iterrows():
                dias_tabla.append([str(r['DIA_SESION']), f"{r['Ocup_Pct']:.2f}%", f"{r['Flujo_Estudiantes']:.2f}"])
            
            t_dias = Table(dias_tabla, colWidths=[120, 110, 110], repeatRows=1)
            t_dias.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#37474f')),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, -1), 8),
                ('GRID', (0, 0), (-1, -1), 0.4, colors.HexColor('#cccccc')),
                ('ALIGN', (1, 0), (2, -1), 'CENTER'),
                ('TOPPADDING', (0, 0), (-1, -1), 3),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
            ]))
            story.append(t_dias)
        else:
            story.append(Paragraph("No hay datos de bloques horarios para días lunes a viernes.", style_normal))
    else:
        story.append(Paragraph("No se dispone de datos de bloques horarios.", style_normal))
    
    story.append(Spacer(1, 12))
    
    # =============================================
    # 7. RECOMENDACIONES (DINÁMICAS)
    # =============================================
    story.append(Paragraph("7. Recomendaciones", style_subtitulo))
    
    recomendaciones = []
    
    # R1: Salas críticas <= 20%
    # salas_criticas ya excluye "Sin Capacidad" desde el paso anterior
    n_criticas_final = len(salas_criticas) if not salas_criticas.empty else 0
    if n_criticas_final > 0:
        salas_top5 = salas_criticas.sort_values('Ocupacion_Pct').head(5)
        salas_nombres = ', '.join([
            f"Sala {r['SALA']} ({r['CAMPUS_AUX']}-{r['EDIFICIO']}, {r['Ocupacion_Pct']:.1f}%)" 
            for _, r in salas_top5.iterrows()])
        recomendaciones.append(
            f"CRÍTICO: Se identificaron {n_criticas_final} salas con ocupación igual o menor al 20%. "
            f"Se recomienda evaluar la reasignación o reconversión de las más críticas: {salas_nombres}. "
            f"Considerar su uso para actividades complementarias, estudios libres o eventos."
        )
    
    # (Recomendación de ocupación promedio general eliminada por solicitud)
    
    # R3: Bloques horarios bajos (excluyendo sab/dom)
    if 'HR_INICIO_AUX_STR' in df_work.columns and 'DIA_SESION' in df_work.columns:
        df_blq_rec = df_work[~df_work['DIA_SESION'].isin(['SABADO', 'DOMINGO'])].copy()
        if not df_blq_rec.empty:
            blq_rec = df_blq_rec.groupby('HR_INICIO_AUX_STR')['INSCRITOS'].sum().reset_index()
            dias_blq_rec = df_blq_rec.groupby('HR_INICIO_AUX_STR')['DIA_SESION'].nunique().reset_index()
            dias_blq_rec.columns = ['HR_INICIO_AUX_STR', 'N_DIAS']
            blq_rec = blq_rec.merge(dias_blq_rec, on='HR_INICIO_AUX_STR', how='left')
            blq_rec['Pct'] = ((blq_rec['INSCRITOS'] / (total_capacidad * blq_rec['N_DIAS'])) * 100).round(2)
            bloques_muy_bajos = blq_rec[blq_rec['Pct'] <= 20]
            if not bloques_muy_bajos.empty:
                horas = ', '.join(bloques_muy_bajos['HR_INICIO_AUX_STR'].tolist())
                recomendaciones.append(
                    f"Los bloques horarios {horas} presentan ocupación igual o menor al 20% (excluyendo "
                    f"sábado y domingo). Se recomienda evaluar si es viable reducir la oferta en estos "
                    f"horarios o incentivar su uso con actividades académicas complementarias."
                )
    
    # R4: Salas bajo promedio
    n_bajo_prom_f = len(df_bajo_prom) if not df_bajo_prom.empty else 0
    total_salas_f = len(tabla_sala)
    if total_salas_f > 0:
        pct_bajo = (n_bajo_prom_f / total_salas_f) * 100
        if pct_bajo > 40:
            recomendaciones.append(
                f"El {pct_bajo:.1f}% de las salas ({n_bajo_prom_f} de {total_salas_f}) estan bajo el "
                f"promedio de su campus. Esto sugiere una distribución desigual y concentración de uso "
                f"en pocas salas."
            )
    
    # R5: Clasificaciones con baja ocupación
    if 'CLASIFICACIÓN' in df_work.columns:
        clasif_ocup = df_work.groupby('CLASIFICACIÓN')[col_ocup].mean().reset_index()
        clasif_ocup['Pct'] = (clasif_ocup[col_ocup] * 100).round(2)
        clasif_bajas = clasif_ocup[clasif_ocup['Pct'] < 15]
        if not clasif_bajas.empty:
            tipos = ', '.join([f"{r['CLASIFICACIÓN']} ({r['Pct']:.1f}%)" for _, r in clasif_bajas.iterrows()])
            recomendaciones.append(
                f"Las siguientes clasificaciones de sala tienen ocupación promedio bajo 15%: {tipos}. "
                f"Se recomienda evaluar si estos tipos de espacio se ajustan a la demanda actual."
            )
    
    # R6: Campus cruzados (CAMPUS != CAMPUS_AUX)
    if 'CAMPUS' in df_work.columns and 'CAMPUS_AUX' in df_work.columns:
        df_cruzados = df_work[df_work['CAMPUS'] != df_work['CAMPUS_AUX']].copy()
        if not df_cruzados.empty:
            n_salas_cruzadas = df_cruzados['SALA'].nunique()
            recomendaciones.append(
                f"ALERTA: Se detectaron {n_salas_cruzadas} salas programadas en un campus diferente al "
                f"campus físico donde se encuentran (CAMPUS de programación difiere de CAMPUS_AUX). "
                f"Se recomienda revisar estas asignaciones para asegurar consistencia entre la "
                f"programación y la ubicación real de los espacios."
            )
    
    # R7: Salas sin capacidad y salas "Sin Clasificación" - buscar en ambas fuentes
    n_sin_cap = 0
    n_sin_clasif = 0
    
    # Buscar desde df3 (planta física)
    if 'Capacidad Máxima Sala' in df3_ref.columns:
        filtro_df3_rec = df3_ref['Campus Edificio'].isin(campus_edificio_list)
        if excl_cap:
            filtro_df3_rec = filtro_df3_rec & (~df3_ref['CLASIFICACIÓN'].isin(excl_cap))
        df3_check = df3_ref[filtro_df3_rec].copy()
        
        n_sin_cap = len(df3_check[
            (df3_check['Capacidad Máxima Sala'].isna()) | (df3_check['Capacidad Máxima Sala'] == 0)
        ])
        
        n_sin_clasif = len(df3_check[
            (df3_check['CLASIFICACIÓN'].isna()) | 
            (df3_check['CLASIFICACIÓN'].astype(str).str.strip() == '') |
            (df3_check['CLASIFICACIÓN'].astype(str).str.strip() == 'Sin Clasificación') |
            (df3_check['CLASIFICACIÓN'].astype(str).str.strip() == 'Sin Clasificacion') |
            (df3_check['CLASIFICACIÓN'].astype(str) == 'nan')
        ])
    
    # Complementar con datos filtrados (Rango_Capacidad = "Sin Capacidad")
    if 'Rango_Capacidad' in df_work.columns:
        salas_sin_cap_work = df_work[df_work['Rango_Capacidad'].astype(str) == 'Sin Capacidad']['SALA'].nunique()
        if salas_sin_cap_work > n_sin_cap:
            n_sin_cap = salas_sin_cap_work
    
    if 'CLASIFICACIÓN' in df_work.columns:
        clasif_vals = ['Sin Clasificación', 'Sin Clasificacion', 'nan', '']
        salas_sin_clasif_work = df_work[
            df_work['CLASIFICACIÓN'].astype(str).str.strip().isin(clasif_vals)
        ]['SALA'].nunique()
        if salas_sin_clasif_work > n_sin_clasif:
            n_sin_clasif = salas_sin_clasif_work
    
    # Siempre agregar esta recomendación (incluso si son 0, para que se vea la verificación)
    texto_datos = "RECTIFICACIÓN DE DATOS: "
    partes = []
    if n_sin_cap > 0:
        partes.append(f"{n_sin_cap} salas no tienen capacidad registrada o tienen capacidad 0")
    if n_sin_clasif > 0:
        partes.append(f"{n_sin_clasif} salas figuran con 'Sin Clasificación' en su clasificación")
    
    if partes:
        texto_datos += ' y '.join(partes) + ". "
        texto_datos += ("Se sugiere rectificar estos datos por el área correspondiente en el sistema "
                       "de planta física, ya que afectan la precisión de los análisis de ocupación.")
        recomendaciones.append(texto_datos)
    
    # Escribir recomendaciones con ancho completo
    if recomendaciones:
        page_width = letter[0] - 3*cm  # ancho disponible
        for i, rec in enumerate(recomendaciones, 1):
            # Usar Paragraph dentro de Table para que el texto haga wrap correctamente
            para = Paragraph(f"<b>{i}.</b> {rec}", 
                           ParagraphStyle('RecInner', parent=styles['Normal'], fontSize=8, leading=11))
            rec_table = Table([[para]], colWidths=[page_width])
            
            bg_color = colors.HexColor('#ffebee') if ('CRÍTICO' in rec or 'ALERTA' in rec or 'DATOS' in rec) else colors.HexColor('#e3f2fd')
            rec_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, -1), bg_color),
                ('BOX', (0, 0), (-1, -1), 0.5, colors.HexColor('#90a4ae')),
                ('LEFTPADDING', (0, 0), (-1, -1), 8),
                ('RIGHTPADDING', (0, 0), (-1, -1), 8),
                ('TOPPADDING', (0, 0), (-1, -1), 5),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
            ]))
            story.append(rec_table)
            story.append(Spacer(1, 4))
    else:
        story.append(Paragraph("No se generaron recomendaciones específicas.", style_normal))
    
    # =============================================
    # PIE DE PÁGINA
    # =============================================
    story.append(Spacer(1, 20))
    story.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor('#bdbdbd')))
    story.append(Paragraph(
        f"Documento generado automáticamente | {fecha_gen} | Total registros analizados: {len(df_filtrado):,}",
        ParagraphStyle('Pie', parent=styles['Normal'], fontSize=7,
                       alignment=TA_CENTER, textColor=colors.HexColor('#999999'))))
    
    # =============================================
    # GENERAR PDF
    # =============================================
    try:
        doc.build(story)
        n_bajo_prom_f = len(df_bajo_prom) if not df_bajo_prom.empty else 0
        print(f"✅ Resumen Ejecutivo generado: {filename}")
        print(f"   📄 Ubicacion: {os.path.abspath(filename)}")
        print(f"   📊 Salas analizadas: {len(tabla_sala)}")
        print(f"   ⚠️ Salas críticas (<=20%): {n_criticas_final}")
        print(f"   📉 Salas bajo promedio campus: {n_bajo_prom_f}")
        print(f"   💡 Recomendaciones generadas: {len(recomendaciones)}")
        return filename
    except Exception as e:
        print(f"❌ Error al generar PDF: {str(e)}")
        import traceback
        traceback.print_exc()
        return None


