import pandas as pd
from openpyxl.styles import PatternFill, Font, Alignment

# =================================================================
# CONFIGURACIÓN DE RUTAS
# =================================================================
archivo_entrada = 'Inventario_desordenado.xlsx'
nombre_salida_excel = 'Inventario_Limpio.xlsx'

try:
    # -------------------------------------------------------------
    # 1. CARGA NATIVA DE EXCEL
    # -------------------------------------------------------------
    # Al usar read_excel, Pandas mantiene la integridad de las columnas
    # y detecta automáticamente los tipos de datos (fechas, números).
    df = pd.read_excel(archivo_entrada)
    # Normalización de encabezados (se cambian a formato de titulo).
    df.columns = df.columns.str.strip()

    # -------------------------------------------------------------
    # 2. ALGORITMOS DE LIMPIEZA Y NORMALIZACIÓN
    # -------------------------------------------------------------
    df['Fecha_Ingreso'] = pd.to_datetime(df['Fecha_Ingreso'], errors='coerce')
    df['Fecha_Ingreso'] = df['Fecha_Ingreso'].dt.strftime('%Y-%m-%d')
    df['Descripcion'] = df['Descripcion'].str.strip().str.title()
    df['ID_Pieza'] = df['ID_Pieza'].str.strip().str.upper()
    df['Cantidad'] = df['Cantidad'].fillna(0)
    df['Fecha_Ingreso'] = df['Fecha_Ingreso'].fillna('No definido')
    df['Precio_Unitario'] = df['Precio_Unitario'].fillna(0)
    # -------------------------------------------------------------
    # 4. EXPORTACIÓN CON DISEÑO PROFESIONAL (OPENPYXL)
    # -------------------------------------------------------------
    with pd.ExcelWriter(nombre_salida_excel, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Inventario')
        worksheet = writer.sheets['Inventario']

        
        # Definición de Estilos (Identidad Visual del Reporte)
        amarillo_relleno = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
        fuente_negrita = Font(bold=True)
        for cell in worksheet[1]:
            cell.fill = amarillo_relleno
            cell.font = fuente_negrita

        # ---------------------------------------------------------
        # 5. POST-PROCESADO: AUTO-AJUSTE Y ALINEACIÓN ESPECÍFICA
        # ---------------------------------------------------------
        for col in worksheet.columns:
            max_length = 0
            column = col[0].column_letter
            for cell in col:
                if cell.column == 5: 
                    cell.alignment = Alignment(horizontal='right')
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except: pass
            # Ajuste de ancho con margen de seguridad.
            worksheet.column_dimensions[column].width = max_length + 2

    print(f"\n--- Reorganizacion del archivo completada con exito. ---")
    print(f"Revisa '{nombre_salida_excel}'.")
except FileNotFoundError:
    print(f"[ERROR] No se encontró el archivo de origen: {archivo_entrada}")
except Exception as e:
    print(f"Hubo un error: {e}")
