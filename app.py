import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Gestión de Atenciones y Cierre Mensual", layout="wide", page_icon="🏥")
st.title("🏥 Gestor de Reportes y Cierre Mensual (27 al 26)")

# Subida de archivo
archivos_subidos = st.file_uploader("Sube uno o varios archivos Excel", type=['xls', 'xlsx'], accept_multiple_files=True)

if archivos_subidos:
    lista_dfs = []
    for archivo in archivos_subidos:
        try:
            try:
                dfs = pd.read_html(archivo, header=1)
                df_temp = dfs[0]
            except:
                df_temp = pd.read_excel(archivo, header=1)
            lista_dfs.append(df_temp)
        except Exception as e:
            st.error(f"Error al leer {archivo.name}: {e}")

    if lista_dfs:
        df_raw = pd.concat(lista_dfs, ignore_index=True)
        
        # Limpieza de datos
        df_limpio = pd.DataFrame()
        col_serie = 'Serie' if 'Serie' in df_raw.columns else df_raw.columns[2]
        col_num = 'Numero' if 'Numero' in df_raw.columns else df_raw.columns[3]
        col_paciente = 'Apellidos y Nombres' if 'Apellidos y Nombres' in df_raw.columns else 'Paciente'
        col_medico = 'Medico' if 'Medico' in df_raw.columns else 'Médico'
        col_desc = 'Producto' if 'Producto' in df_raw.columns else 'Descripción'
        col_monto = 'Subtotal' if 'Subtotal' in df_raw.columns else 'Total'
        col_fecha = 'Fecha Atencion' if 'Fecha Atencion' in df_raw.columns else 'Fecha'
        
        df_limpio['Nº TICKET'] = df_raw[col_serie].astype(str) + "-" + df_raw[col_num].astype(str)
        df_limpio['PACIENTE'] = df_raw[col_paciente]
        df_limpio['MEDICO'] = df_raw[col_medico]
        df_limpio['DESCRIPCION'] = df_raw[col_desc]
        df_limpio['MONTO'] = pd.to_numeric(df_raw[col_monto].astype(str).str.replace(',', '').str.replace('S/', ''), errors='coerce')
        
        # Procesamiento de Fechas (día primero)
        df_limpio['FECHA_DT'] = pd.to_datetime(df_raw[col_fecha], dayfirst=True, errors='coerce')
        df_limpio['FECHA'] = df_limpio['FECHA_DT'].dt.date
        
        # --- LÓGICA DE CIERRE MENSUAL (27 AL 26) ---
        def calcular_periodo(fecha):
            if pd.isna(fecha):
                return "Sin Fecha"
            # Si el día es 27 o más, corresponde al cierre del próximo mes
            if fecha.day >= 27:
                mes_cierre = fecha.month + 1 if fecha.month < 12 else 1
                año_cierre = fecha.year if fecha.month < 12 else fecha.year + 1
            else:
                mes_cierre = fecha.month
                año_cierre = fecha.year
                
            nombres_meses = {1:'Enero', 2:'Febrero', 3:'Marzo', 4:'Abril', 5:'Mayo', 6:'Junio', 
                             7:'Julio', 8:'Agosto', 9:'Septiembre', 10:'Octubre', 11:'Noviembre', 12:'Diciembre'}
            return f"Cierre {nombres_meses[mes_cierre]} {año_cierre}"

        df_limpio['PERIODO'] = df_limpio['FECHA_DT'].apply(calcular_periodo)
        df_limpio = df_limpio.dropna(subset=['PACIENTE', 'MONTO'])
        
        st.divider()
        
        # --- FILTRO POR PERIODO ---
        periodos_disponibles = sorted(df_limpio['PERIODO'].unique().tolist())
        periodo_seleccionado = st.selectbox("📅 Selecciona el Periodo de Cierre que deseas analizar:", periodos_disponibles)
        
        # Aplicar el filtro a la tabla
        df_filtrado = df_limpio[df_limpio['PERIODO'] == periodo_seleccionado].copy()
        
        # Eliminar columnas auxiliares para la vista final
        df_exportar = df_filtrado.drop(columns=['FECHA_DT', 'PERIODO'])

        # --- INTERFAZ CON PESTAÑAS ---
        tab1, tab2 = st.tabs(["📄 Registros Filtrados", "📊 Cierre Mensual"])
        
        with tab1:
            st.dataframe(df_exportar, use_container_width=True)
            
            output_simple = BytesIO()
            with pd.ExcelWriter(output_simple, engine='xlsxwriter') as writer:
                df_exportar.to_excel(writer, index=False, sheet_name='Reporte Limpio')
            
            st.download_button("⬇️ Descargar Tabla de este Cierre", data=output_simple.getvalue(), file_name=f"Reporte_{periodo_seleccionado}.xlsx", mime="application/vnd.ms-excel")

        with tab2:
            total_ingreso = df_filtrado['MONTO'].sum()
            total_pacientes = df_filtrado['PACIENTE'].nunique()
            total_atenciones = len(df_filtrado)
            
            col1, col2, col3 = st.columns(3)
            col1.metric("💰 Ingreso del Periodo", f"S/ {total_ingreso:,.2f}")
            col2.metric("👥 Pacientes Únicos", total_pacientes)
            col3.metric("🩺 Total de Tickets", total_atenciones)
            
            st.divider()
            col_graf_1, col_graf_2 = st.columns(2)
            
            with col_graf_1:
                st.write("**Ingresos por Médico**")
                resumen_medico = df_filtrado.groupby('MEDICO').agg(
                    Total_Ingresos=('MONTO', 'sum'),
                    N_Atenciones=('Nº TICKET', 'count')
                ).reset_index().sort_values('Total_Ingresos', ascending=False)
                st.dataframe(resumen_medico.style.format({'Total_Ingresos': 'S/ {:,.2f}'}), use_container_width=True)
                
            with col_graf_2:
                st.write("**Ingresos por Día**")
                resumen_fecha = df_filtrado.groupby('FECHA').agg(Total_Ingreso=('MONTO', 'sum')).reset_index().sort_values('FECHA')
                st.bar_chart(resumen_fecha.set_index('FECHA'))
            
            st.divider()
            
            output_cierre = BytesIO()
            with pd.ExcelWriter(output_cierre, engine='xlsxwriter') as writer:
                df_exportar.to_excel(writer, index=False, sheet_name='Atenciones')
                resumen_medico.to_excel(writer, index=False, sheet_name='Por Medico')
                resumen_fecha.to_excel(writer, index=False, sheet_name='Por Dia')
            
            st.download_button(label="📊 Descargar Cierre Contable Completo", data=output_cierre.getvalue(), file_name=f"Cierre_Contable_{periodo_seleccionado}.xlsx", mime="application/vnd.ms-excel", type='primary')
