"""
app.py - Dashboard de Ocupación Académica
Ejecutar: python app.py
Acceder: http://localhost:8050
"""
import dash
from dash import dcc, html, Input, Output, State, callback, no_update, ctx
import dash_bootstrap_components as dbc
import plotly.graph_objects as go
import plotly.express as px
import pandas as pd
import numpy as np
import os
import tempfile
from pathlib import Path
from datetime import date

from utils.data_loader import cargar_datos, obtener_opciones_filtros
from utils.pdf_generator import generar_resumen_ejecutivo_pdf, get_download_folder

# ============================================================
# CARGA DE DATOS
# ============================================================
print("📂 Cargando datos...")
df_aperturado, df2, df3 = cargar_datos('BBDD')
opciones = obtener_opciones_filtros(df_aperturado, df3)
print(f"✅ Datos cargados: {len(df_aperturado):,} registros")

# ============================================================
# APP DASH
# ============================================================
app = dash.Dash(
    __name__,
    suppress_callback_exceptions=True,
    external_stylesheets=[
        'https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600;700&family=Playfair+Display:wght@400;600;700&display=swap'
    ],
    title='Ocupación Académica'
)

# ============================================================
# LAYOUT
# ============================================================

def make_filter_label(text):
    return html.Label(text, className='filter-label')

# Sidebar
sidebar = html.Div(className='sidebar', children=[
    # Header
    html.Div(className='sidebar-header', children=[
        html.Div(className='logo-area', children=[
            html.Div('OA', className='logo-icon'),
            html.Div('Ocupación Académica', className='logo-title'),
        ]),
        html.Div('Panel de análisis', className='sidebar-subtitle'),
    ]),
    
    # Filtros generales
    html.Div(className='sidebar-section', children=[
        html.Div('Filtros generales', className='section-label'),
        
        make_filter_label('Año'),
        dcc.Dropdown(
            id='filtro-anio',
            options=[{'label': str(a), 'value': a} for a in opciones['anios']],
            value=2025 if 2025 in opciones['anios'] else opciones['anios'][-1] if opciones['anios'] else None,
            clearable=False,
        ),
        
        make_filter_label('Campus Edificio'),
        dcc.Dropdown(
            id='filtro-campus-edificio',
            options=[{'label': str(c), 'value': c} for c in opciones['campus_edificios']],
            value=[opciones['campus_edificios'][0]] if opciones['campus_edificios'] else [],
            multi=True,
        ),
        
        make_filter_label('Campus Programación'),
        dcc.Dropdown(
            id='filtro-campus',
            options=[{'label': str(c), 'value': c} for c in opciones['campus_list']],
            value=[opciones['campus_list'][0]] if opciones['campus_list'] else [],
            multi=True,
        ),
    ]),
    
    # Programación
    html.Div(className='sidebar-section', children=[
        html.Div('Programación', className='section-label'),
        
        make_filter_label('Período Académico'),
        dcc.Dropdown(
            id='filtro-periodo',
            options=[{'label': str(p), 'value': p} for p in opciones['periodos']],
            value=[opciones['periodos'][0]] if opciones['periodos'] else [],
            multi=True,
        ),
        
        make_filter_label('Carrera / Materia'),
        dcc.Dropdown(
            id='filtro-carrera',
            options=[{'label': str(c), 'value': c} for c in opciones['desc_carreras']],
            value=[],
            multi=True,
            placeholder='Todas',
        ),
        
        html.Div(style={'marginTop': '10px'}, children=[
            dcc.Checklist(
                id='chk-fechas',
                options=[{'label': ' Filtrar por rango de fechas', 'value': 'usar'}],
                value=[],
                style={'fontSize': '12px'},
            ),
        ]),
        
        html.Div(id='div-fechas', style={'opacity': '0.5'}, children=[
            make_filter_label('Fecha inicio'),
            dcc.DatePickerSingle(
                id='filtro-fecha-ini',
                date=opciones['fecha_min'],
                display_format='DD/MM/YYYY',
                style={'width': '100%'},
            ),
            make_filter_label('Fecha fin'),
            dcc.DatePickerSingle(
                id='filtro-fecha-fin',
                date=opciones['fecha_max'],
                display_format='DD/MM/YYYY',
                style={'width': '100%'},
            ),
        ]),
    ]),
    
    # Clasificaciones
    html.Div(className='sidebar-section', children=[
        html.Div('Clasificaciones a excluir', className='section-label'),
        
        make_filter_label('Excluir en Capacidad'),
        dcc.Dropdown(
            id='filtro-excl-cap',
            options=[{'label': str(c), 'value': c} for c in opciones['clasificaciones_capacidad']],
            value=['EXTERNO'] if 'EXTERNO' in opciones['clasificaciones_capacidad'] else [],
            multi=True,
        ),
        
        make_filter_label('Excluir en Programación'),
        dcc.Dropdown(
            id='filtro-excl-ocu',
            options=[{'label': str(c), 'value': c} for c in opciones['clasificaciones_ocupacion']],
            value=[],
            multi=True,
        ),
        
        # Filtros extra Sunburst
        html.Div(id='sunburst-extra', className='sunburst-extra-box', children=[
            html.Span('Solo Sunburst', className='badge-sunburst'),
            
            make_filter_label('Bloque inicio'),
            dcc.Dropdown(
                id='filtro-bloque-ini',
                options=[{'label': b, 'value': b} for b in opciones['bloques_inicio']],
                value=None,
                placeholder='Todos',
            ),
            
            make_filter_label('Bloque fin'),
            dcc.Dropdown(
                id='filtro-bloque-fin',
                options=[{'label': b, 'value': b} for b in opciones['bloques_fin']],
                value=None,
                placeholder='Todos',
            ),
            
            make_filter_label('Días sesión'),
            dcc.Dropdown(
                id='filtro-dias',
                options=[{'label': d, 'value': d} for d in opciones['dias_disponibles']],
                value=opciones['dias_disponibles'],
                multi=True,
            ),
        ]),
    ]),
    
    # Botones
    html.Div(className='sidebar-actions', children=[
        html.Button('⟳  Actualizar', id='btn-actualizar', className='btn-actualizar', n_clicks=0),
        html.Button('↓  Descargar CSV', id='btn-csv', className='btn-csv', n_clicks=0),
        html.Button('⬇  Resumen Ejecutivo PDF', id='btn-pdf', className='btn-pdf', n_clicks=0),
    ]),
])

# Main content
main_content = html.Div(className='main-content', children=[
    # Top bar
    html.Div(className='top-bar', children=[
        html.H2('Dashboard de Ocupación', className='page-title'),
        html.Span(id='timestamp-display', className='timestamp-badge'),
    ]),
    
    # KPI cards
    html.Div(id='kpi-row', className='kpi-row'),
    
    # Tabs
    dcc.Tabs(id='tabs', value='heatmap', className='custom-tabs', children=[
        dcc.Tab(label='Mapa de Calor', value='heatmap', className='custom-tab', selected_className='custom-tab--selected'),
        dcc.Tab(label='Sunburst', value='sunburst', className='custom-tab', selected_className='custom-tab--selected'),
    ]),
    
    # Chart
    html.Div(className='chart-box', children=[
        dcc.Loading(
            dcc.Graph(id='main-chart', config={'displayModeBar': True, 'locale': 'es'}),
            type='circle',
            color='#c4956a',
        ),
    ]),
    
    # Summary
    html.Div(id='summary-row', className='summary-row'),
    
    # Hidden stores
    dcc.Store(id='store-df-filtrado'),
    dcc.Store(id='store-filtros'),
    dcc.Store(id='store-total-capacidad'),
    dcc.Download(id='download-csv'),
    dcc.Download(id='download-pdf'),
    html.Div(id='msg-output', style={'display': 'none'}),
])

# App layout
# Panel de personalización
config_panel = html.Div(id='config-overlay', className='config-overlay config-hidden', children=[
    html.Div(className='config-panel', children=[
        html.Div(className='config-header', children=[
            html.H3('⚙️ Personalización', style={'margin': '0', 'fontFamily': 'Playfair Display, serif'}),
            html.Button('✕', id='btn-config-close', className='config-close-btn'),
        ]),
        
        html.Div(className='config-body', children=[
            # Transparencia del fondo
            html.Div(className='config-group', children=[
                html.Label('Transparencia del fondo', className='config-label'),
                dcc.Slider(
                    id='slider-opacity',
                    min=0, max=30, step=1, value=6,
                    marks={0: '0%', 6: '6%', 10: '10%', 15: '15%', 20: '20%', 30: '30%'},
                    tooltip={'placement': 'bottom', 'always_visible': False},
                ),
            ]),
            
            # Subir imagen de fondo
            html.Div(className='config-group', children=[
                html.Label('Imagen de fondo', className='config-label'),
                html.Div(className='config-hint', children='Sube una imagen PNG, JPG o WEBP'),
                dcc.Upload(
                    id='upload-bg',
                    children=html.Div(['📁 Arrastra o ', html.A('selecciona archivo', style={'color': 'var(--accent-copper)', 'cursor': 'pointer'})]),
                    className='config-upload',
                    accept='.png,.jpg,.jpeg,.webp',
                ),
                html.Button('↩ Restaurar fondo original', id='btn-reset-bg', className='config-reset-btn'),
            ]),
            
            # Color de acento
            html.Div(className='config-group', children=[
                html.Label('Color de acento', className='config-label'),
                html.Div(className='config-color-row', children=[
                    html.Button(className='color-swatch', id='swatch-copper',
                                style={'background': 'linear-gradient(135deg, #c4956a, #b8734d)'},
                                title='Cobre (original)'),
                    html.Button(className='color-swatch', id='swatch-blue',
                                style={'background': 'linear-gradient(135deg, #5b8fb9, #3a6d94)'},
                                title='Azul acero'),
                    html.Button(className='color-swatch', id='swatch-sage',
                                style={'background': 'linear-gradient(135deg, #8a9a7b, #6b7d5e)'},
                                title='Verde sage'),
                    html.Button(className='color-swatch', id='swatch-plum',
                                style={'background': 'linear-gradient(135deg, #9b7a9d, #7d5a7f)'},
                                title='Ciruela'),
                    html.Button(className='color-swatch', id='swatch-slate',
                                style={'background': 'linear-gradient(135deg, #6b7b8d, #4a5a6b)'},
                                title='Pizarra'),
                ]),
            ]),
            
            # Tema sidebar
            html.Div(className='config-group', children=[
                html.Label('Tema del sidebar', className='config-label'),
                dcc.RadioItems(
                    id='radio-sidebar-theme',
                    options=[
                        {'label': ' Claro', 'value': 'light'},
                        {'label': ' Oscuro', 'value': 'dark'},
                    ],
                    value='light',
                    inline=True,
                    style={'fontSize': '13px'},
                ),
            ]),
            
            # Paleta heatmap
            html.Div(className='config-group', children=[
                html.Label('Paleta del Heatmap', className='config-label'),
                dcc.Dropdown(
                    id='select-heatmap-palette',
                    options=[
                        {'label': 'YlOrRd (original)', 'value': 'YlOrRd'},
                        {'label': 'RdYlGn', 'value': 'RdYlGn'},
                        {'label': 'Viridis', 'value': 'Viridis'},
                        {'label': 'Plasma', 'value': 'Plasma'},
                        {'label': 'Inferno', 'value': 'Inferno'},
                        {'label': 'Blues', 'value': 'Blues'},
                        {'label': 'Turbo', 'value': 'Turbo'},
                    ],
                    value='YlOrRd',
                    clearable=False,
                ),
            ]),
            
            # Paleta sunburst
            html.Div(className='config-group', children=[
                html.Label('Paleta del Sunburst', className='config-label'),
                dcc.Dropdown(
                    id='select-sunburst-palette',
                    options=[
                        {'label': 'RdYlGn (original)', 'value': 'RdYlGn'},
                        {'label': 'YlOrRd', 'value': 'YlOrRd'},
                        {'label': 'Viridis', 'value': 'Viridis'},
                        {'label': 'Plasma', 'value': 'Plasma'},
                        {'label': 'Spectral', 'value': 'Spectral'},
                        {'label': 'Turbo', 'value': 'Turbo'},
                    ],
                    value='RdYlGn',
                    clearable=False,
                ),
            ]),
        ]),
    ]),
])

# Stores para personalización
config_stores = html.Div([
    dcc.Store(id='store-heatmap-palette', data='YlOrRd'),
    dcc.Store(id='store-sunburst-palette', data='RdYlGn'),
    dcc.Store(id='store-bg-image', data=None),
    dcc.Store(id='store-accent-color', data='copper'),
])

# Botón flotante de configuración
config_button = html.Button('⚙️', id='btn-config-open', className='config-fab')

# CSS dinámico se inyecta via JS (clientside callback)

# App layout
app.layout = html.Div(className='app-container', children=[
    sidebar, main_content, config_panel, config_button, config_stores
])


# ============================================================
# CALLBACKS
# ============================================================

# --- Activar/desactivar fechas ---
@app.callback(
    Output('div-fechas', 'style'),
    Input('chk-fechas', 'value'),
)
def toggle_fechas(chk):
    if 'usar' in (chk or []):
        return {'opacity': '1'}
    return {'opacity': '0.5', 'pointerEvents': 'none'}


# --- Mostrar/ocultar filtros sunburst según tab ---
@app.callback(
    Output('sunburst-extra', 'style'),
    Input('tabs', 'value'),
)
def toggle_sunburst_filters(tab):
    if tab == 'sunburst':
        return {'display': 'block'}
    return {'display': 'none'}


# --- Timestamp ---
@app.callback(
    Output('timestamp-display', 'children'),
    Input('btn-actualizar', 'n_clicks'),
)
def update_timestamp(_):
    return pd.Timestamp.now().strftime('%d / %b / %Y — %H:%M')


# ============================================================
# CALLBACK PRINCIPAL: Generar gráfico + KPIs + summary
# ============================================================
@app.callback(
    Output('main-chart', 'figure'),
    Output('kpi-row', 'children'),
    Output('summary-row', 'children'),
    Output('store-df-filtrado', 'data'),
    Output('store-filtros', 'data'),
    Output('store-total-capacidad', 'data'),
    Input('btn-actualizar', 'n_clicks'),
    Input('tabs', 'value'),
    State('filtro-anio', 'value'),
    State('filtro-campus-edificio', 'value'),
    State('filtro-campus', 'value'),
    State('filtro-periodo', 'value'),
    State('filtro-carrera', 'value'),
    State('chk-fechas', 'value'),
    State('filtro-fecha-ini', 'date'),
    State('filtro-fecha-fin', 'date'),
    State('filtro-excl-cap', 'value'),
    State('filtro-excl-ocu', 'value'),
    State('filtro-bloque-ini', 'value'),
    State('filtro-bloque-fin', 'value'),
    State('filtro-dias', 'value'),
    Input('store-heatmap-palette', 'data'),
    Input('store-sunburst-palette', 'data'),
)
def actualizar_grafico(n_clicks, tab, anio, campus_edificio_list, campus_list,
                       periodo_list, carrera_list, chk_fechas, fecha_ini, fecha_fin,
                       excl_cap, excl_ocu, bloque_ini, bloque_fin, dias_list,
                       heatmap_palette, sunburst_palette):
    
    campus_edificio_list = campus_edificio_list or []
    campus_list = campus_list or []
    periodo_list = periodo_list or []
    carrera_list = carrera_list or []
    excl_cap = excl_cap or []
    excl_ocu = excl_ocu or []
    dias_list = dias_list or []
    
    usar_fechas = 'usar' in (chk_fechas or [])
    fecha_inicio = fecha_ini if usar_fechas else None
    fecha_fin_val = fecha_fin if usar_fechas else None
    
    # Filtros dict para PDF
    filtros_dict = {
        'anio': anio,
        'campus_list': campus_list,
        'campus_edificio_list': campus_edificio_list,
        'periodo_list': periodo_list,
        'desc_carrera_materia_list': carrera_list,
        'fecha_inicio': fecha_inicio,
        'fecha_fin': fecha_fin_val,
        'clasificacion_excluir_capacidad': excl_cap,
        'clasificacion_excluir_ocupacion': excl_ocu,
    }
    
    # Capacidad total
    filtro_df3 = df3['Campus Edificio'].isin(campus_edificio_list)
    if excl_cap:
        filtro_df3 = filtro_df3 & (~df3['CLASIFICACIÓN'].isin(excl_cap))
    total_capacidad = df3.loc[filtro_df3, 'Capacidad Máxima Sala'].sum()
    if total_capacidad == 0:
        fig = go.Figure()
        fig.add_annotation(text="⚠️ No hay datos de capacidad con los filtros seleccionados", showarrow=False, font=dict(size=16))
        return fig, [], [], None, filtros_dict, 0
    
    # Filtrar ocupación
    df_ocupacion = df_aperturado[
        (df_aperturado['REG_UNICO_LLAV2'] == 1) &
        (df_aperturado['DIA_SESION'].notna()) &
        (df_aperturado['DIA_SESION'] != '') &
        (df_aperturado['SALA'] != '')
    ].copy()
    
    df_ocupacion = df_ocupacion[
        (df_ocupacion['AÑO'] == anio) &
        (df_ocupacion['CAMPUS'].isin(campus_list))
    ]
    
    if periodo_list:
        df_ocupacion = df_ocupacion[df_ocupacion['PERIODO'].isin(periodo_list)]
    if carrera_list:
        df_ocupacion = df_ocupacion[df_ocupacion['DESCRIPCIÓN CARRERA_MATERIA'].isin(carrera_list)]
    
    if fecha_inicio and fecha_fin_val:
        if 'FECHA_INI2' in df_ocupacion.columns and 'FECHA_TERM2' in df_ocupacion.columns:
            df_ocupacion['FECHA_INI2'] = pd.to_datetime(df_ocupacion['FECHA_INI2'], origin='1899-12-30', unit='D', errors='coerce')
            df_ocupacion['FECHA_TERM2'] = pd.to_datetime(df_ocupacion['FECHA_TERM2'], origin='1899-12-30', unit='D', errors='coerce')
            fecha_inicio_dt = pd.to_datetime(fecha_inicio)
            fecha_fin_dt = pd.to_datetime(fecha_fin_val)
            df_ocupacion = df_ocupacion[
                (df_ocupacion['FECHA_INI2'] <= fecha_fin_dt) &
                (df_ocupacion['FECHA_TERM2'] >= fecha_inicio_dt)
            ]
            df_ocupacion['FECHA_INI_AUX'] = df_ocupacion['FECHA_INI2'].apply(
                lambda x: x if pd.notna(x) and x >= fecha_inicio_dt else fecha_inicio_dt)
            df_ocupacion['FECHA_TERM_AUX'] = df_ocupacion['FECHA_TERM2'].apply(
                lambda x: x if pd.notna(x) and x <= fecha_fin_dt else fecha_fin_dt)
    
    if excl_ocu:
        df_ocupacion = df_ocupacion[~df_ocupacion['CLASIFICACIÓN'].isin(excl_ocu)]
    
    if df_ocupacion.empty:
        fig = go.Figure()
        fig.add_annotation(text="⚠️ No hay datos de ocupación con los filtros seleccionados", showarrow=False, font=dict(size=16))
        return fig, [], [], None, filtros_dict, total_capacidad
    
    # ---- GENERAR GRÁFICO SEGÚN TAB ----
    if tab == 'heatmap':
        fig, info = generar_heatmap(df_ocupacion, df2, total_capacidad, filtros_dict, heatmap_palette or 'YlOrRd')
    else:
        # Filtros extra sunburst
        if dias_list:
            df_ocupacion = df_ocupacion[df_ocupacion['DIA_SESION'].isin(dias_list)]
        if bloque_ini and bloque_fin:
            df_ocupacion['HR_INICIO_AUX_STR'] = df_ocupacion['HR_INICIO_AUX_STR'].astype(str)
            df_ocupacion['HR_FIN_AUX_STR'] = df_ocupacion['HR_FIN_AUX_STR'].astype(str)
            df_ocupacion = df_ocupacion[
                (df_ocupacion['HR_INICIO_AUX_STR'] >= bloque_ini) &
                (df_ocupacion['HR_FIN_AUX_STR'] <= bloque_fin)
            ]
        
        filtros_dict['bloque_inicio'] = bloque_ini
        filtros_dict['bloque_fin'] = bloque_fin
        filtros_dict['dias_sesion_list'] = dias_list
        
        fig, info = generar_sunburst(df_ocupacion, df3, filtros_dict, sunburst_palette or 'RdYlGn')
    
    # KPIs
    df_valid = df_ocupacion[
        (df_ocupacion['SALA'].notna()) & (df_ocupacion['SALA'] != '') & (df_ocupacion['SALA'].astype(str) != 'nan')
    ]
    ocup_prom = (df_valid['OCUPACIÓN_SALA'].mean() * 100) if 'OCUPACIÓN_SALA' in df_valid.columns and not df_valid.empty else 0
    n_salas = df_valid['SALA'].nunique() if not df_valid.empty else 0
    
    # Salas críticas (promedio por sala <= 20%)
    if not df_valid.empty and 'OCUPACIÓN_SALA' in df_valid.columns:
        sala_prom = df_valid.groupby('SALA')['OCUPACIÓN_SALA'].mean() * 100
        n_criticas = (sala_prom <= 20).sum()
    else:
        n_criticas = 0
    
    kpis = [
        html.Div(className='kpi-card', children=[
            html.Div('Ocupación Promedio', className='kpi-label'),
            html.Div(f'{ocup_prom:.1f}%', className='kpi-value'),
            html.Div(f'Campus: {", ".join(campus_list[:2])}{"..." if len(campus_list) > 2 else ""}', className='kpi-sub'),
        ]),
        html.Div(className='kpi-card', children=[
            html.Div('Salas Analizadas', className='kpi-label'),
            html.Div(f'{n_salas}', className='kpi-value'),
            html.Div(f'{len(campus_edificio_list)} edificios', className='kpi-sub'),
        ]),
        html.Div(className='kpi-card', children=[
            html.Div('Salas Críticas', className='kpi-label'),
            html.Div(f'{n_criticas}', className='kpi-value red'),
            html.Div('Ocupación ≤ 20%', className='kpi-sub'),
        ]),
        html.Div(className='kpi-card', children=[
            html.Div('Capacidad Total', className='kpi-label'),
            html.Div(f'{total_capacidad:,.0f}', className='kpi-value'),
            html.Div('Butacas disponibles', className='kpi-sub'),
        ]),
    ]
    
    # Summary
    periodos_str = ', '.join([str(p) for p in periodo_list]) if periodo_list else 'Todos'
    summary = [
        html.Div(className='summary-item', children=[
            html.Div('Registros filtrados', className='summary-label'),
            html.Div(f'{len(df_ocupacion):,}', className='summary-value'),
        ]),
        html.Div(className='summary-item', children=[
            html.Div('Edificios seleccionados', className='summary-label'),
            html.Div(f'{len(campus_edificio_list)}', className='summary-value'),
        ]),
        html.Div(className='summary-item', children=[
            html.Div('Período', className='summary-label'),
            html.Div(periodos_str, className='summary-value'),
        ]),
    ]
    
    # Guardar df filtrado como JSON para descargas
    df_store = df_ocupacion.head(50000).to_json(date_format='iso', orient='split')
    
    return fig, kpis, summary, df_store, filtros_dict, total_capacidad


# ============================================================
# FUNCIONES DE GRÁFICO
# ============================================================

def generar_heatmap(df_ocupacion, df2, total_capacidad, filtros_dict, colorscale='YlOrRd'):
    """Genera el heatmap de ocupación."""
    orden_dias = ['LUNES', 'MARTES', 'MIERCOLES', 'JUEVES', 'VIERNES', 'SABADO', 'DOMINGO']
    
    df2_capacidad = df2.replace('X', total_capacidad)
    df_ocupacion['HR_INICIO_AUX_STR'] = df_ocupacion['HR_INICIO_AUX_STR'].astype(str)
    
    matriz_ocupacion = pd.pivot_table(
        df_ocupacion, values='INSCRITOS', index='HR_INICIO_AUX_STR',
        columns='DIA_SESION', aggfunc='sum', fill_value=0
    )
    
    for dia in orden_dias:
        if dia not in matriz_ocupacion.columns:
            matriz_ocupacion[dia] = 0
    matriz_ocupacion = matriz_ocupacion[orden_dias]
    
    df2_capacidad['HORA_INICIO_STR'] = df2_capacidad['HORA INICIO'].astype(str)
    
    df_cap_melted = df2_capacidad.melt(
        id_vars=['BLOQUE', 'HORA_INICIO_STR'], value_vars=orden_dias,
        var_name='DIA', value_name='CAPACIDAD_TOTAL'
    )
    
    matriz_reset = matriz_ocupacion.reset_index().rename(columns={'HR_INICIO_AUX_STR': 'HORA_INICIO_STR'})
    ocu_melted = matriz_reset.melt(
        id_vars=['HORA_INICIO_STR'], value_vars=orden_dias,
        var_name='DIA', value_name='OCUPACION'
    )
    
    df_comb = pd.merge(df_cap_melted, ocu_melted, on=['HORA_INICIO_STR', 'DIA'], how='left')
    df_comb['OCUPACION'] = df_comb['OCUPACION'].fillna(0)
    df_comb['PORCENTAJE'] = (df_comb['OCUPACION'] / df_comb['CAPACIDAD_TOTAL']) * 100
    
    df_heatmap = df_comb.pivot_table(
        index=['BLOQUE', 'HORA_INICIO_STR'], columns='DIA', values='PORCENTAJE', aggfunc='first'
    )[orden_dias].reset_index()
    
    horas_dict = df2_capacidad.set_index('HORA_INICIO_STR')[['HORA INICIO', 'HORA FIN']].to_dict('index')
    df_heatmap['HORA INICIO'] = df_heatmap['HORA_INICIO_STR'].map(lambda x: horas_dict.get(x, {}).get('HORA INICIO', ''))
    df_heatmap['HORA FIN'] = df_heatmap['HORA_INICIO_STR'].map(lambda x: horas_dict.get(x, {}).get('HORA FIN', ''))
    df_heatmap = df_heatmap.sort_values('BLOQUE').reset_index(drop=True)
    
    z = df_heatmap[orden_dias].values
    bloque_labels = 'B' + df_heatmap['BLOQUE'].astype(str) + ' ' + \
                    df_heatmap['HORA INICIO'].astype(str) + '-' + df_heatmap['HORA FIN'].astype(str)
    
    fig = go.Figure(data=go.Heatmap(
        z=z, x=orden_dias, y=bloque_labels,
        colorscale=colorscale, zmin=0, zmax=100,
        text=[[f'{val:.2f}%' for val in row] for row in z],
        texttemplate='%{text}', textfont={"size": 10},
        hovertemplate='<b>Día:</b> %{x}<br><b>Bloque:</b> %{y}<br><b>Ocupación:</b> %{z:.2f}%<extra></extra>',
        colorbar=dict(title="Ocupación (%)", ticksuffix="%")
    ))
    
    campus_str = ', '.join(filtros_dict.get('campus_list', [])[:3])
    edificio_str = ', '.join(filtros_dict.get('campus_edificio_list', [])[:3])
    
    fig.update_layout(
        title=f'Mapa de Calor — Ocupación Académica de Campus<br><sub>Campus: {campus_str} | Edificios: {edificio_str} | Año: {filtros_dict.get("anio", "")}</sub>',
        xaxis_title='Día de la Semana', yaxis_title='Bloque Horario',
        yaxis={'autorange': 'reversed'},
        height=700, font=dict(family='DM Sans', size=12),
        paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
        margin=dict(l=20, r=20, t=80, b=20),
    )
    
    return fig, {'total_capacidad': total_capacidad, 'registros': len(df_ocupacion)}


def generar_sunburst(df_ocupacion, df3, filtros_dict, colorscale='RdYlGn'):
    """Genera el gráfico sunburst."""
    campus_edificio_list = filtros_dict.get('campus_edificio_list', [])
    excl_cap = filtros_dict.get('clasificacion_excluir_capacidad', [])
    
    # Capacidades por sala
    filtro_df3 = df3['Campus Edificio'].isin(campus_edificio_list)
    if excl_cap:
        filtro_df3 = filtro_df3 & (~df3['CLASIFICACIÓN'].isin(excl_cap))
    filtro_df3 = filtro_df3 & (df3['Campus Edificio'] != 'ONL') & (df3['Campus Edificio'] != 0)
    
    df3_f = df3[filtro_df3].copy()
    if df3_f.empty:
        fig = go.Figure()
        fig.add_annotation(text="⚠️ Sin datos de capacidad", showarrow=False)
        return fig, {}
    
    df_cap_sala = df3_f.groupby(['Campus Edificio', 'Código Edificio', 'Número Sala']).agg(
        {'Capacidad Máxima Sala': 'sum'}).reset_index()
    df_cap_sala.columns = ['CAMPUS_AUX', 'EDIFICIO', 'SALA', 'CAPACIDAD_SALA']
    for c in ['CAMPUS_AUX', 'EDIFICIO', 'SALA']:
        df_cap_sala[c] = df_cap_sala[c].astype(str)
    
    # Ocupación
    df_ocu = df_ocupacion.copy()
    for c in ['CAMPUS_AUX', 'EDIFICIO', 'SALA']:
        if c in df_ocu.columns:
            df_ocu[c] = df_ocu[c].astype(str)
    
    df_ocu = df_ocu[df_ocu['CAMPUS_AUX'] != 'ONL']
    if campus_edificio_list:
        df_ocu = df_ocu[df_ocu['CAMPUS_AUX'].isin([str(c) for c in campus_edificio_list])]
    
    if df_ocu.empty or 'Rango_Capacidad' not in df_ocu.columns:
        fig = go.Figure()
        fig.add_annotation(text="⚠️ Sin datos de ocupación", showarrow=False)
        return fig, {}
    
    # Métricas por sala
    metricas_sala = df_ocu.groupby(['CAMPUS_AUX', 'EDIFICIO', 'Rango_Capacidad', 'SALA']).agg(
        {'OCUPACIÓN_SALA': 'mean', 'INSCRITOS': ['sum', 'max', 'min']}).reset_index()
    metricas_sala.columns = ['CAMPUS_AUX', 'EDIFICIO', 'Rango_Capacidad', 'SALA', 'TIP', 'INS_TOT', 'INS_MAX', 'INS_MIN']
    
    metricas_sala = metricas_sala.merge(
        df_cap_sala[['CAMPUS_AUX', 'EDIFICIO', 'SALA', 'CAPACIDAD_SALA']],
        on=['CAMPUS_AUX', 'EDIFICIO', 'SALA'], how='left')
    metricas_sala['CAPACIDAD_SALA'] = metricas_sala['CAPACIDAD_SALA'].fillna(30)
    metricas_sala['VALUE'] = metricas_sala['CAPACIDAD_SALA']
    
    # Niveles agregados
    met_rango = df_ocu.groupby(['CAMPUS_AUX', 'EDIFICIO', 'Rango_Capacidad']).agg(
        {'OCUPACIÓN_SALA': 'mean', 'SALA': 'nunique'}).reset_index()
    met_rango.columns = ['CAMPUS_AUX', 'EDIFICIO', 'Rango_Capacidad', 'TIP', 'N_SALAS']
    val_rango = metricas_sala.groupby(['CAMPUS_AUX', 'EDIFICIO', 'Rango_Capacidad'])['VALUE'].sum().reset_index()
    met_rango = met_rango.merge(val_rango, on=['CAMPUS_AUX', 'EDIFICIO', 'Rango_Capacidad'], how='left')
    met_rango['VALUE'] = met_rango['VALUE'].fillna(1)
    
    met_edif = df_ocu.groupby(['CAMPUS_AUX', 'EDIFICIO']).agg({'OCUPACIÓN_SALA': 'mean'}).reset_index()
    met_edif.columns = ['CAMPUS_AUX', 'EDIFICIO', 'TIP']
    val_edif = met_rango.groupby(['CAMPUS_AUX', 'EDIFICIO'])['VALUE'].sum().reset_index()
    met_edif = met_edif.merge(val_edif, on=['CAMPUS_AUX', 'EDIFICIO'], how='left')
    met_edif['VALUE'] = met_edif['VALUE'].fillna(1)
    
    met_campus = df_ocu.groupby('CAMPUS_AUX').agg({'OCUPACIÓN_SALA': 'mean'}).reset_index()
    met_campus.columns = ['CAMPUS_AUX', 'TIP']
    val_campus = met_edif.groupby('CAMPUS_AUX')['VALUE'].sum().reset_index()
    met_campus = met_campus.merge(val_campus, on='CAMPUS_AUX', how='left')
    met_campus['VALUE'] = met_campus['VALUE'].fillna(1)
    
    # Construir sunburst
    ids, labels, parents, values, clrs, hovers = [], [], [], [], [], []
    
    for _, r in met_campus.iterrows():
        c = str(r['CAMPUS_AUX'])
        ids.append(c); labels.append(c); parents.append('')
        values.append(float(r['VALUE'])); clrs.append(float(r['TIP']) if pd.notna(r['TIP']) else 0)
        hovers.append(f"<b>{c}</b><br>Ocupación: {(r['TIP']*100) if pd.notna(r['TIP']) else 0:.2f}%")
    
    for _, r in met_edif.iterrows():
        c, e = str(r['CAMPUS_AUX']), str(r['EDIFICIO'])
        id_e = f"{c} - {e}"
        ids.append(id_e); labels.append(e); parents.append(c)
        values.append(float(r['VALUE'])); clrs.append(float(r['TIP']) if pd.notna(r['TIP']) else 0)
        hovers.append(f"<b>{e}</b><br>Ocupación: {(r['TIP']*100) if pd.notna(r['TIP']) else 0:.2f}%")
    
    for _, r in met_rango.iterrows():
        c, e, rg = str(r['CAMPUS_AUX']), str(r['EDIFICIO']), str(r['Rango_Capacidad'])
        id_r = f"{c} - {e} - {rg}"
        ids.append(id_r); labels.append(rg); parents.append(f"{c} - {e}")
        values.append(float(r['VALUE'])); clrs.append(float(r['TIP']) if pd.notna(r['TIP']) else 0)
        hovers.append(f"<b>{rg}</b><br>Ocupación: {(r['TIP']*100) if pd.notna(r['TIP']) else 0:.2f}%<br>Salas: {r['N_SALAS']:.0f}")
    
    for _, r in metricas_sala.iterrows():
        c, e, rg, s = str(r['CAMPUS_AUX']), str(r['EDIFICIO']), str(r['Rango_Capacidad']), str(r['SALA'])
        id_s = f"{c} - {e} - {rg} - {s}"
        ids.append(id_s); labels.append(f"Sala {s}"); parents.append(f"{c} - {e} - {rg}")
        values.append(float(r['VALUE'])); clrs.append(float(r['TIP']) if pd.notna(r['TIP']) else 0)
        hovers.append(f"<b>Sala {s}</b><br>Capacidad: {r['CAPACIDAD_SALA']:,.0f}<br>Ocupación: {(r['TIP']*100) if pd.notna(r['TIP']) else 0:.2f}%")
    
    fig = go.Figure(go.Sunburst(
        ids=ids, labels=labels, parents=parents, values=values,
        marker=dict(colors=clrs, colorscale=colorscale, cmin=0, cmax=1,
                    colorbar=dict(title="Ocupación", tickformat=".0%", tickvals=[0, 0.25, 0.5, 0.75, 1],
                                  ticktext=['0%', '25%', '50%', '75%', '100%'])),
        hovertext=hovers, hoverinfo='text', branchvalues='total'
    ))
    
    campus_str = ', '.join(filtros_dict.get('campus_list', []))
    fig.update_layout(
        title=dict(
            text=f'<b>Ocupación por Campus — Vista Sunburst</b><br><sub>Campus: {campus_str} | Año: {filtros_dict.get("anio", "")}</sub>',
            x=0.5, xanchor='center'),
        height=800, font=dict(family='DM Sans', size=12),
        paper_bgcolor='rgba(0,0,0,0)', margin=dict(t=100, l=20, r=20, b=20),
    )
    
    return fig, {}


# ============================================================
# DESCARGA CSV
# ============================================================
@app.callback(
    Output('download-csv', 'data'),
    Input('btn-csv', 'n_clicks'),
    State('store-df-filtrado', 'data'),
    prevent_initial_call=True,
)
def descargar_csv(n_clicks, df_json):
    if not df_json:
        return no_update
    df = pd.read_json(df_json, orient='split')
    timestamp = pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')
    return dcc.send_data_frame(df.to_csv, f'datos_filtrados_{timestamp}.csv', index=False, encoding='utf-8-sig')


# ============================================================
# DESCARGA PDF
# ============================================================
@app.callback(
    Output('download-pdf', 'data'),
    Input('btn-pdf', 'n_clicks'),
    State('store-df-filtrado', 'data'),
    State('store-filtros', 'data'),
    State('store-total-capacidad', 'data'),
    State('main-chart', 'figure'),
    prevent_initial_call=True,
)
def descargar_pdf(n_clicks, df_json, filtros_dict, total_cap, fig_dict):
    if not df_json or not filtros_dict:
        return no_update
    
    df = pd.read_json(df_json, orient='split')
    
    # Reconstruir figura para exportar imagen
    fig_img = None
    if fig_dict:
        try:
            fig_img = go.Figure(fig_dict)
        except:
            pass
    
    filepath = generar_resumen_ejecutivo_pdf(
        df_filtrado=df,
        df3_ref=df3,
        df2_ref=df2,
        filtros_dict=filtros_dict,
        tipo_grafico='Dashboard',
        fig_imagen=fig_img,
    )
    
    if filepath and os.path.exists(filepath):
        return dcc.send_file(filepath)
    return no_update


# ============================================================
# PERSONALIZACIÓN - CALLBACKS
# ============================================================

# --- Abrir/cerrar panel ---
app.clientside_callback(
    """
    function(n_open, n_close) {
        const overlay = document.getElementById('config-overlay');
        if (!overlay) return dash_clientside.no_update;
        
        const triggered = dash_clientside.callback_context.triggered;
        if (!triggered || triggered.length === 0) return dash_clientside.no_update;
        
        const prop_id = triggered[0].prop_id;
        if (prop_id === 'btn-config-open.n_clicks') {
            overlay.classList.remove('config-hidden');
        } else {
            overlay.classList.add('config-hidden');
        }
        return dash_clientside.no_update;
    }
    """,
    Output('config-overlay', 'className'),
    Input('btn-config-open', 'n_clicks'),
    Input('btn-config-close', 'n_clicks'),
    prevent_initial_call=True,
)


# --- Personalización via clientside callback (JS puro, sin html.Style) ---
app.clientside_callback(
    """
    function(opacity, n_cop, n_blu, n_sag, n_plu, n_sla,
             sidebar_theme, bg_contents, n_reset,
             stored_bg, stored_accent) {
        
        // Detectar qué disparó el callback
        var triggered = '';
        try {
            var ctx = dash_clientside.callback_context;
            if (ctx && ctx.triggered && ctx.triggered.length > 0) {
                triggered = ctx.triggered[0].prop_id.split('.')[0];
            }
        } catch(e) {}
        
        // Color de acento
        var colorMap = {
            'swatch-copper': ['#c4956a', '#b8734d'],
            'swatch-blue':   ['#5b8fb9', '#3a6d94'],
            'swatch-sage':   ['#8a9a7b', '#6b7d5e'],
            'swatch-plum':   ['#9b7a9d', '#7d5a7f'],
            'swatch-slate':  ['#6b7b8d', '#4a5a6b']
        };
        
        var accentKey = stored_accent || 'copper';
        var keyMap = {
            'swatch-copper': 'copper', 'swatch-blue': 'blue',
            'swatch-sage': 'sage', 'swatch-plum': 'plum', 'swatch-slate': 'slate'
        };
        if (keyMap[triggered]) { accentKey = keyMap[triggered]; }
        
        var accentWarm = '#c4956a', accentCopper = '#b8734d';
        var accentLookup = {
            'copper': ['#c4956a','#b8734d'], 'blue': ['#5b8fb9','#3a6d94'],
            'sage': ['#8a9a7b','#6b7d5e'], 'plum': ['#9b7a9d','#7d5a7f'],
            'slate': ['#6b7b8d','#4a5a6b']
        };
        if (accentLookup[accentKey]) {
            accentWarm = accentLookup[accentKey][0];
            accentCopper = accentLookup[accentKey][1];
        }
        
        // Opacity
        var op = (opacity || 6) / 100.0;
        
        // Sidebar theme
        var sbBg, sbText, sbBorder, sbLabel;
        if (sidebar_theme === 'dark') {
            sbBg='#2d2a27'; sbText='#e0dbd5'; sbBorder='#3d3a37'; sbLabel='#a09a94';
        } else {
            sbBg='#ffffff'; sbText='#3d3a37'; sbBorder='#e0dbd5'; sbLabel='#918b85';
        }
        
        // Background image
        var currentBg = stored_bg;
        if (triggered === 'btn-reset-bg') { currentBg = null; }
        else if (triggered === 'upload-bg' && bg_contents) { currentBg = bg_contents; }
        
        var bgUrl = currentBg || "/assets/marble_bg.webp";
        
        // Inyectar CSS
        var css = `
            body::before { opacity: ${op} !important; background-image: url('${bgUrl}') !important; }
            :root { --accent-warm: ${accentWarm} !important; --accent-copper: ${accentCopper} !important; }
            .sidebar { background: ${sbBg} !important; border-right-color: ${sbBorder} !important; }
            .sidebar .section-label, .sidebar .filter-label { color: ${sbLabel} !important; }
            .sidebar .logo-title, .sidebar-section { color: ${sbText} !important; }
            .sidebar-header { background: ${sbBg} !important; border-bottom-color: ${sbBorder} !important; }
            .sidebar-section { border-bottom-color: ${sbBorder} !important; }
            .logo-icon { background: linear-gradient(135deg, ${accentWarm}, ${accentCopper}) !important;
                         box-shadow: 0 2px 8px ${accentCopper}4d !important; }
            .btn-pdf { background: linear-gradient(135deg, ${accentWarm}, ${accentCopper}) !important; }
            .custom-tab--selected { border-bottom-color: ${accentCopper} !important; }
            .kpi-card:nth-child(1)::before { background: ${accentWarm} !important; }
            .kpi-card:nth-child(4)::before { background: ${accentCopper} !important; }
        `;
        
        // Inyectar o actualizar tag <style> en el DOM
        var styleEl = document.getElementById('dynamic-injected-style');
        if (!styleEl) {
            styleEl = document.createElement('style');
            styleEl.id = 'dynamic-injected-style';
            document.head.appendChild(styleEl);
        }
        styleEl.textContent = css;
        
        return [currentBg, accentKey];
    }
    """,
    Output('store-bg-image', 'data'),
    Output('store-accent-color', 'data'),
    Input('slider-opacity', 'value'),
    Input('swatch-copper', 'n_clicks'),
    Input('swatch-blue', 'n_clicks'),
    Input('swatch-sage', 'n_clicks'),
    Input('swatch-plum', 'n_clicks'),
    Input('swatch-slate', 'n_clicks'),
    Input('radio-sidebar-theme', 'value'),
    Input('upload-bg', 'contents'),
    Input('btn-reset-bg', 'n_clicks'),
    State('store-bg-image', 'data'),
    State('store-accent-color', 'data'),
    prevent_initial_call=True,
)


# --- Guardar paletas en stores ---
@app.callback(
    Output('store-heatmap-palette', 'data'),
    Input('select-heatmap-palette', 'value'),
)
def update_heatmap_palette(val):
    return val or 'YlOrRd'

@app.callback(
    Output('store-sunburst-palette', 'data'),
    Input('select-sunburst-palette', 'value'),
)
def update_sunburst_palette(val):
    return val or 'RdYlGn'


# ============================================================
# RUN
# ============================================================
if __name__ == '__main__':
    print("\n🚀 Dashboard disponible en: http://localhost:8050\n")
    app.run(debug=True, host='0.0.0.0', port=8050)
