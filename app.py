import streamlit as st
import pandas as pd
from datetime import datetime
import unicodedata
import io

# --- CONFIGURACIÓN VISUAL DE LA PÁGINA ---
st.set_page_config(page_title="Reporte Ocupación TPL", page_icon="🚍", layout="centered")

# --- DICCIONARIO DE RUTAS ---
RUTAS_TPL = {
    '001': ['Terminal Varoli', 'Terminal Talca', 'Casa Matriz', '2 norte con carretera', 'Cruce Pelarco', 'Cruce San Rafael', 'Cruce Camarico', 'Rancagua Sur', 'Colon San Bernardo', 'Terminal Sur'],
    '002': ['Terminal Sur', 'Colon San Bernardo', 'Cruce Rengo', 'Terminal Varoli', 'Terminal Talca'],
    '003': ['Terminal Varoli', 'Terminal Talca', 'Casa matriz', 'Cruce Varoli', 'Cruce Maule', 'Cruce San Javier Sur', 'Cruce Villa Alegre', 'Cruce Linares (Petrobras)', 'Cruce Longavi', 'Cruce Parral (Hospital)', 'Cruce Copihue', 'Cruce San Gregorio', 'Terminal María Teresa', 'Plaza Chillán Viejo', 'KM 21 (Nueva Aldea)', 'Terminal Collao'],
    '004': ['Terminal Collao', 'El Manzano', 'Enlace Penco', 'KM 21 (Nueva Aldea)', 'KM 21(Nueva Aldea).', 'Terminal María Teresa', 'San Carlos La Virgen', 'Cruce San Javier', 'Terminal Varoli', 'Terminal Talca'],
    '005': ['Terminal Varoli', 'Terminal Talca', 'Casa matriz', 'Cruce Varoli', 'Cruce Maule', 'Cruce San Javier Sur', 'Cruce Villa Alegre', 'Cruce Linares (Petrobras)', 'Cruce Longavi', 'Cruce Parral (Hospital)', 'Cruce Copihue', 'Cruce San Gregorio', 'Terminal María Teresa'],
    '006': ['Terminal María Teresa', 'San Carlos La Virgen', 'Cruce San Javier', 'Terminal Varoli', 'Terminal Talca'],
    '007': ['Terminal Cauquenes', 'San Javier Frente PDI', 'Cruce San Javier', 'Terminal Varoli', 'Terminal Talca', '2 norte con carretera', 'Cruce Pelarco', 'Cruce San Rafael', 'Cruce Camarico', 'Rancagua Sur', 'Colon San Bernardo', 'Terminal Sur'],
    '008': ['Terminal Sur', 'Colon San Bernardo', 'Cruce Rengo', 'Terminal Varoli', 'Terminal Talca', 'Cruce Varoli', 'San Javier Frente PDI', 'Terminal Cauquenes'],
    '009': ['Terminal Varoli', 'Terminal Talca', 'Casa matriz', 'Cruce Varoli', 'Cruce Maule', 'Cruce San Javier Sur', 'Cruce Villa Alegre', 'Cruce Linares (Petrobras)', 'Cruce Longavi', 'Cruce Parral (Hospital)', 'Cruce Copihue', 'Cruce San Gregorio', 'Terminal María Teresa', 'Plaza Chillán Viejo', 'Cruce Bulnes', 'Cruce Cabrero', 'Cruce Laja', 'Cruce Perales', 'Terminal María Teresa'], 
    '010': ['Terminal María Teresa', 'Cruce Perales', 'Cruce Laja', 'Cruce Cabrero', 'Cruce Bulnes', 'Terminal María Teresa', 'San Carlos La Virgen', 'Cruce San Javier', 'Terminal Varoli', 'Terminal Talca'],
    '011': ['Frente Municipalidad', 'Paradero Cesfam', 'Terminal Cauquenes', 'San Javier Frente PDI', 'Cruce San Javier', 'Terminal Varoli', 'Terminal Talca', '2 norte con carretera', 'Cruce Pelarco', 'Cruce San Rafael', 'Cruce Camarico', 'Rancagua Sur', 'Colon San Bernardo', 'Terminal Sur'],
    '012': ['Terminal Sur', 'Colon San Bernardo', 'Cruce Rengo', 'Terminal Varoli', 'Terminal Talca', 'Cruce Varoli', 'San Javier Frente PDI', 'Terminal Cauquenes', 'Paradero Cesfam', 'Frente Municipalidad']
}

# --- FUNCIONES ---
def normalizar_texto(texto):
    if not isinstance(texto, str): return ""
    texto = texto.strip().lower()
    return ''.join(c for c in unicodedata.normalize('NFKD', texto) if not unicodedata.combining(c))

def parse_date_time(date_str, time_str):
    try:
        if isinstance(date_str, datetime): d_str = date_str.strftime('%d-%m-%Y')
        else: d_str = str(date_str).strip().replace('/', '-')
        if isinstance(time_str, datetime): t_str = time_str.strftime('%H:%M')
        else: 
            t_str = str(time_str).strip()
            if len(t_str) == 8: t_str = t_str[:5]
        return pd.to_datetime(f"{d_str} {t_str}", format='%d-%m-%Y %H:%M', errors='coerce')
    except: return pd.NaT

def get_stop_order(ruta_str, paradero):
    if not isinstance(ruta_str, str): return 999
    ruta_num = str(ruta_str).split()[0]
    if ruta_num in RUTAS_TPL:
        stops = [normalizar_texto(s) for s in RUTAS_TPL[ruta_num]]
        p_clean = normalizar_texto(paradero)
        if p_clean in stops: return stops.index(p_clean)
        for i, s in enumerate(stops):
            if p_clean in s or s in p_clean: return i
    return 999

# --- INTERFAZ WEB ---
st.title("🚍 Generador de Reporte de Ocupación")
st.markdown("Sube tu archivo **Reporte resumen de ventas** original y obtén al instante el reporte de ocupación ordenado.")

# Cajón para arrastrar el archivo
archivo_subido = st.file_uploader("Arrastra tu Excel aquí", type=["xlsx"])

if archivo_subido is not None:
    with st.spinner("Procesando el archivo de forma automática..."):
        try:
            xls = pd.ExcelFile(archivo_subido)
            sheet_name = xls.sheet_names[0]
            df = pd.read_excel(xls, sheet_name=sheet_name)
            
            if 'Estado' not in df.columns:
                st.error("❌ El archivo no parece ser el reporte correcto (falta la columna 'Estado').")
            else:
                df_pagados = df[df['Estado'] == 'Pagado'].copy()
                df_pagados = df_pagados.dropna(subset=['Folio'])
                df_pagados['Folio'] = df_pagados['Folio'].astype(int).astype(str)
                
                df_pagados['DateTime_Origen'] = df_pagados.apply(lambda row: parse_date_time(row['Fecha salida'], row['Hora salida']), axis=1)
                df_pagados['DateTime_Destino'] = df_pagados.apply(lambda row: parse_date_time(row['Fecha llegada'], row['Hora llegada']), axis=1)
                
                paradas_list = []
                for _, row in df_pagados.iterrows():
                    folio, ruta = row['Folio'], row['Ruta']
                    paradas_list.append({
                        'Folio': folio, 'Ruta': ruta,
                        'Ciudad': str(row['Origen (ciudad)']).strip() if pd.notna(row['Origen (ciudad)']) else '',
                        'Paradero': str(row['Origen (parada)']).strip() if pd.notna(row['Origen (parada)']) else '',
                        'DateTime': row['DateTime_Origen'],
                    })
                    paradas_list.append({
                        'Folio': folio, 'Ruta': ruta,
                        'Ciudad': str(row['Destino (ciudad)']).strip() if pd.notna(row['Destino (ciudad)']) else '',
                        'Paradero': str(row['Destino (parada)']).strip() if pd.notna(row['Destino (parada)']) else '',
                        'DateTime': row['DateTime_Destino'],
                    })

                df_paradas = pd.DataFrame(paradas_list).drop_duplicates(subset=['Folio', 'Ciudad', 'Paradero']).copy()
                df_paradas['Orden_Oficial'] = df_paradas.apply(lambda row: get_stop_order(row['Ruta'], row['Paradero']), axis=1)

                res_dfs = []
                for folio, group in df_paradas.groupby('Folio'):
                    group = group.sort_values(by=['DateTime', 'Orden_Oficial']).reset_index(drop=True)
                    for i, row in group.iterrows():
                        c, p = row['Ciudad'], row['Paradero']
                        sub = df_pagados[(df_pagados['Folio'] == folio) & (df_pagados['Origen (ciudad)'].str.strip() == c) & (df_pagados['Origen (parada)'].str.strip() == p)].shape[0]
                        baj = df_pagados[(df_pagados['Folio'] == folio) & (df_pagados['Destino (ciudad)'].str.strip() == c) & (df_pagados['Destino (parada)'].str.strip() == p)].shape[0]
                        group.at[i, 'Suben'] = sub
                        group.at[i, 'Bajan'] = baj
                        group.at[i, 'Revisar orden'] = 'Sí' if row['Orden_Oficial'] == 999 else ''
                    res_dfs.append(group)
                
                if not res_dfs:
                    st.warning("⚠️ No se encontraron pasajes válidos para procesar.")
                else:
                    df_final = pd.concat(res_dfs, ignore_index=True)
                    for folio, group in df_final.groupby('Folio', sort=False):
                        idx_list = group.index.tolist()
                        for j, i in enumerate(idx_list):
                            if j == 0: df_final.at[i, 'Continúan'] = 0
                            else: df_final.at[i, 'Continúan'] = df_final.at[idx_list[j-1], 'Total'] - df_final.at[i, 'Bajan']
                            df_final.at[i, 'Total'] = df_final.at[i, 'Continúan'] + df_final.at[i, 'Suben']
                    
                    dias_semana = {0: 'LUNES', 1: 'MARTES', 2: 'MIÉRCOLES', 3: 'JUEVES', 4: 'VIERNES', 5: 'SÁBADO', 6: 'DOMINGO'}
                    df_final['Día'] = df_final['DateTime'].dt.dayofweek.map(dias_semana)
                    df_final['Fecha salida'] = df_final['DateTime'].dt.strftime('%d/%m/%Y')
                    df_final['Hora'] = df_final['DateTime'].dt.strftime('%H:%M')
                    
                    cols = ['Folio', 'Fecha salida', 'Día', 'Ciudad', 'Hora', 'Paradero', 'Bajan', 'Suben', 'Continúan', 'Total', 'Revisar orden']
                    df_final = df_final[cols]
                    
                    # Preparar el Excel en memoria para descarga
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df_final.to_excel(writer, sheet_name='Ocupación', index=False)
                    
                    st.success("✅ ¡Reporte generado con éxito!")
                    
                    # Mostrar una muestra en pantalla
                    st.write("👀 **Vista previa de los datos:**")
                    st.dataframe(df_final.head(10)) 
                    
                    # Botón gigante de descarga
                    st.download_button(
                        label="📥 Descargar Ocupación (Excel)",
                        data=output.getvalue(),
                        file_name=archivo_subido.name.replace('.xlsx', '_ocupacion.xlsx'),
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
        except Exception as e:
            st.error(f"❌ Ocurrió un error inesperado: {e}")
