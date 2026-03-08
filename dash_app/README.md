# Dashboard de Ocupación Académica

## Requisitos previos

```bash
pip install dash dash-bootstrap-components plotly pandas numpy reportlab kaleido
```

## Estructura de archivos

```
dash_app/
├── app.py                      # App principal (ejecutar este)
├── README.md                   # Este archivo
├── assets/
│   ├── style.css              # Estilos (aquí cambias opacity del fondo)
│   └── marble_bg.webp         # Textura de mármol
├── utils/
│   ├── __init__.py
│   ├── data_loader.py         # Carga de CSVs
│   └── pdf_generator.py       # Generador de PDF
└── BBDD/                       # ← Colocar aquí los 3 CSV
    ├── PROG_DETALLADA_PARA_HEATMAP.csv
    ├── BLOQUES_ESTANDAR_01122025.csv
    └── PLANTA_FISICA_PREPROCESADA.csv
```

## Cómo ejecutar

1. Copiar los 3 archivos CSV en la carpeta `BBDD/`
2. Ejecutar:
   ```bash
   cd dash_app
   python app.py
   ```
3. Abrir en el navegador: **http://localhost:8050**

## Personalización

### Cambiar intensidad del fondo marmoleado
Editar `assets/style.css`, buscar `opacity: 0.06` y cambiar:
- `0.03` → Apenas visible
- `0.06` → Sutil (actual)
- `0.10` → Visible
- `0.15` → Decorativo

### Puerto
Cambiar en `app.py` la línea: `app.run(debug=True, host='0.0.0.0', port=8050)`

### Acceso desde otros equipos en la red
La app ya escucha en `0.0.0.0`, así que cualquier equipo en la misma red
puede acceder usando la IP del servidor: `http://IP_DEL_SERVIDOR:8050`
