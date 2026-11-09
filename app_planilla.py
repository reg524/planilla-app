import datetime
import io
import xml.etree.ElementTree as ET
import pandas as pd
import requests
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
import streamlit as st
from streamlit_gsheets import GSheetsConnection

# Configuración de página en Streamlit
st.set_page_config(
    page_title="Sistema de Planilla Bimoneda, Multiempresa y Cargas Sociales",
    page_icon="⚖️",
    layout="wide",
)

# Porcentajes según normativa de Costa Rica
PORCENTAJE_CCSS_OBRERO = 0.1067   # 10.67% CCSS Trabajador
PORCENTAJE_CCSS_PATRONAL = 0.2650  # 26.50% Cargas Sociales Patronales
PORCENTAJE_AGUINALDO = 1 / 12       # 8.33% Aguinaldo
PORCENTAJE_VACACIONES = 1 / 24      # 4.16% Vacaciones
PORCENTAJE_CESANTIA = 0.0533        # 5.33% Cesantía

MESES_MAP = {
    "Enero": 1, "Febrero": 2, "Marzo": 3, "Abril": 4, "Mayo": 5, "Junio": 6,
    "Julio": 7, "Agosto": 8, "Setiembre": 9, "Octubre": 10, "Noviembre": 11, "Diciembre": 12
}
MESES = list(MESES_MAP.keys())

COLUMNAS_HISTORICO = [
    "Empresa", "Año", "Mes", "Cedula", "Nombre", "Puesto", "Fecha_Ingreso", "Fecha_Salida",
    "Moneda", "Tipo_Cambio", "Salario_Ingresado", "Salario_Mensual_CRC", "Salario_Mensual_USD",
    "CCSS_Obrero_CRC", "Salario_Neto_CRC", "CCSS_Patronal_CRC",
    "Aguinaldo_CRC", "Vacaciones_CRC", "Cesantia_CRC", "Total_Prestaciones_CRC"
]

# --- CONEXIÓN DIRECTA A GOOGLE SHEETS EN LA NUBE ---

@st.cache_resource
def get_gsheet_connection():
    return st.connection("gsheets", type=GSheetsConnection)

def cargar_datos_gsheet(worksheet_name, columnas_defecto):
    try:
        conn = get_gsheet_connection()
        df = conn.read(worksheet=worksheet_name, ttl=5)
        for col in columnas_defecto:
            if col not in df.columns:
                df[col] = 0.0 if "CRC" in col or "USD" in col else ""
        return df
    except Exception:
        return pd.DataFrame(columns=columnas_defecto)

def guardar_datos_gsheet(df, worksheet_name):
    try:
        conn = get_gsheet_connection()
        conn.update(worksheet=worksheet_name, data=df)
        st.cache_data.clear()
        return True
    except Exception as e:
        st.error(f"Error al guardar en Google Sheets: {e}")
        return False


def asegurar_calculos_historico(df):
    """Garantiza que todos los rubros de CCSS y Prestaciones existan y estén calculados"""
    if df.empty:
        return df

    if "Salario_Mensual_CRC" not in df.columns:
        df["Salario_Mensual_CRC"] = 0.0

    df["Salario_Mensual_CRC"] = pd.to_numeric(df["Salario_Mensual_CRC"], errors="coerce").fillna(0.0)
    df["Tipo_Cambio"] = pd.to_numeric(df["Tipo_Cambio"], errors="coerce").fillna(515.0)
    df["Tipo_Cambio"] = df["Tipo_Cambio"].apply(lambda x: x if x > 0 else 515.0)

    df["CCSS_Obrero_CRC"] = df["Salario_Mensual_CRC"] * PORCENTAJE_CCSS_OBRERO
    df["Salario_Neto_CRC"] = df["Salario_Mensual_CRC"] - df["CCSS_Obrero_CRC"]
    df["CCSS_Patronal_CRC"] = df["Salario_Mensual_CRC"] * PORCENTAJE_CCSS_PATRONAL
    df["Aguinaldo_CRC"] = df["Salario_Mensual_CRC"] * PORCENTAJE_AGUINALDO
    df["Vacaciones_CRC"] = df["Salario_Mensual_CRC"] * PORCENTAJE_VACACIONES
    df["Cesantia_CRC"] = df["Salario_Mensual_CRC"] * PORCENTAJE_CESANTIA
    df["Total_Prestaciones_CRC"] = df["Aguinaldo_CRC"] + df["Vacaciones_CRC"] + df["Cesantia_CRC"]

    return df


def obtener_tipo_cambio_compra_bccr(fecha_dt):
    """Consulta del Tipo de Cambio de COMPRA BCCR (Indicador 317)"""
    fecha_str = fecha_dt.strftime("%d/%m/%Y")
    url = "https://gee.bccr.fi.cr/Indicadores/Suscripciones/WS/wsindicadoreseconomicos.asmx/ObtenerIndicadoresEconomicosXML"

    params = {
        "Indicador": "317",
        "FechaInicio": fecha_str,
        "FechaFinal": fecha_str,
        "Nombre": "Consulta",
        "SubNiveles": "N",
        "CorreoElectronico": "consultas@bccr.fi.cr",
        "Token": "BCCR",
    }

    try:
        response = requests.get(url, params=params, timeout=4)
        if response.status_code == 200:
            root = ET.fromstring(response.text)
            inner_xml_text = root.text
            if inner_xml_text:
                inner_root = ET.fromstring(inner_xml_text)
                val_node = inner_root.find(".//NUM_VALUE")
                if val_node is not None and val_node.text:
                    return float(val_node.text.replace(",", "."))
    except Exception:
        pass
    return 515.00


# --- GENERACIÓN DE REPORTES ---

def generar_excel_empresa(df_emp, nombre_empresa, elaborado_por, revisado_por):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_emp.to_excel(writer, sheet_name=nombre_empresa[:30], index=False)
    return output.getvalue()


def generar_word_empresa(df_emp, nombre_empresa, elaborado_por, revisado_por):
    doc = Document()
    p_title = doc.add_heading(f"Reporte de Planilla, CCSS y Prestaciones\n{nombre_empresa}", 0)
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    table = doc.add_table(rows=1, cols=len(df_emp.columns))
    table.style = "Table Grid"

    hdr_cells = table.rows[0].cells
    for i, col_name in enumerate(df_emp.columns):
        hdr_cells[i].text = col_name

    for _, row in df_emp.iterrows():
        row_cells = table.add_row().cells
        for i, val in enumerate(row):
            if isinstance(val, float):
                row_cells[i].text = f"₡{val:,.2f}"
            else:
                row_cells[i].text = str(val) if pd.notnull(val) else ""

    doc.add_paragraph("\n")
    p_firmas = doc.add_paragraph()
    p_firmas.add_run(f"Elaborado por: {elaborado_por}\t\t\t\tRevisado por: {revisado_por}")

    output = io.BytesIO()
    doc.save(output)
    return output.getvalue()


def generar_pdf_empresa(df_emp, nombre_empresa, titulo_reporte, elaborado_por, revisado_por):
    output = io.BytesIO()
    doc = SimpleDocTemplate(
        output,
        pagesize=landscape(letter),
        rightMargin=15,
        leftMargin=15,
        topMargin=25,
        bottomMargin=25,
    )
    elements = []
    styles = getSampleStyleSheet()

    style_title = ParagraphStyle(
        'TitleStyle',
        parent=styles['Heading1'],
        fontName='Helvetica-Bold',
        fontSize=13,
        alignment=1,
        spaceAfter=8
    )
    style_cell = ParagraphStyle('CellText', fontName='Helvetica', fontSize=6.5, leading=8, alignment=1)
    style_header = ParagraphStyle('HeaderStyle', fontName='Helvetica-Bold', fontSize=6.5, leading=8, textColor=colors.whitesmoke, alignment=1)

    elements.append(Paragraph(f"<b>{nombre_empresa.upper()}</b>", style_title))
    elements.append(Paragraph(f"<b>{titulo_reporte}</b>", style_title))
    elements.append(Spacer(1, 8))

    headers = [Paragraph(str(col).replace("_", " "), style_header) for col in df_emp.columns]
    data = [headers]

    for _, row in df_emp.iterrows():
        fila_formateada = []
        for col_name, val in zip(df_emp.columns, row):
            if isinstance(val, float):
                texto = f"${val:,.2f}" if "USD" in col_name else f"₡{val:,.2f}"
            else:
                texto = str(val) if pd.notnull(val) else ""
            fila_formateada.append(Paragraph(texto, style_cell))
        data.append(fila_formateada)

    num_cols = len(df_emp.columns)
    col_width = (762) / num_cols

    t = Table(data, colWidths=[col_width] * num_cols)
    t.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F4E79")),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#D9D9D9")),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ]
        )
    )
    elements.append(t)
    elements.append(Spacer(1, 25))

    firmas_data = [
        [
            Paragraph(f"<b>Elaborado por:</b> {elaborado_por}", ParagraphStyle('Firma', fontName='Helvetica', fontSize=8.5)),
            Paragraph(f"<b>Revisado por:</b> {revisado_por}", ParagraphStyle('Firma', fontName='Helvetica', fontSize=8.5)),
        ]
    ]
    t_firmas = Table(firmas_data, colWidths=[381, 381])
    t_firmas.setStyle(TableStyle([("LINEABOVE", (0, 0), (-1, -1), 0.5, colors.black), ("TOPPADDING", (0, 0), (-1, -1), 5)]))
    elements.append(t_firmas)

    doc.build(elements)
    return output.getvalue()


# --- BARRA LATERAL ---
st.sidebar.title("Módulos del Sistema (Sincronizado)")

menu = st.sidebar.radio(
    "Seleccione el Menú:",
    [
        "1. Catálogo de Empresas y Empleados",
        "2. Salarios a Ingresar (Planilla)",
        "3. Costo CCSS Patronal por Empresa",
        "4. Gestor de Registros Históricos",
        "5. Reportes y Exportación",
    ],
)

# --- MÓDULO 1: CATÁLOGO DE EMPRESAS Y EMPLEADOS ---
if menu == "1. Catálogo de Empresas y Empleados":
    st.title("1. Administración de Empresas y Colaboradores")
    st.caption("Los datos guardados aquí se sincronizan inmediatamente en Google Sheets.")

    df_emp = cargar_datos_gsheet("Empresas", ["Empresa", "Cedula_Juridica", "Contacto"])
    df_empld = cargar_datos_gsheet("Empleados", ["Empresa", "Cedula", "Nombre", "Puesto", "Fecha_Ingreso", "Fecha_Salida"])

    tab1, tab2 = st.tabs(["🏢 Catálogo de Empresas", "👤 Padrón de Empleados por Empresa"])

    with tab1:
        st.subheader("Base de Datos de Empresas (Nube)")
        df_emp_edited = st.data_editor(df_emp, num_rows="dynamic", use_container_width=True, key="editor_empresas")
        if st.button("💾 Sincronizar Catálogo de Empresas"):
            if guardar_datos_gsheet(df_emp_edited, "Empresas"):
                st.success("✅ Catálogo de Empresas actualizado y sincronizado.")
                st.rerun()

    with tab2:
        st.subheader("Base de Datos de Empleados (Nube)")
        empresas_list = sorted(df_emp_edited["Empresa"].dropna().unique().tolist())
        
        if not empresas_list:
            st.warning("Primero registre al menos una empresa en el catálogo.")
        else:
            emp_sel_padron = st.selectbox("Seleccione la Empresa:", empresas_list)
            df_sub_empld = df_empld[df_empld["Empresa"] == emp_sel_padron].copy()

            df_sub_edited = st.data_editor(
                df_sub_empld,
                num_rows="dynamic",
                use_container_width=True,
                key=f"editor_empld_{emp_sel_padron}"
            )

            if st.button(f"💾 Sincronizar Padrón de {emp_sel_padron}"):
                df_otros = df_empld[df_empld["Empresa"] != emp_sel_padron]
                df_sub_edited["Empresa"] = emp_sel_padron
                df_final_empld = pd.concat([df_otros, df_sub_edited], ignore_index=True)
                if guardar_datos_gsheet(df_final_empld, "Empleados"):
                    st.success(f"✅ Personal de **{emp_sel_padron}** sincronizado en la nube.")
                    st.rerun()

# --- MÓDULO 2: SALARIOS A INGRESAR (PLANILLA) ---
elif menu == "2. Salarios a Ingresar (Planilla)":
    st.title("2. Salarios a Ingresar")
    st.caption("Alimentación de salarios con sincronización remota instantánea.")

    df_emp = cargar_datos_gsheet("Empresas", ["Empresa", "Cedula_Juridica", "Contacto"])
    df_empld = cargar_datos_gsheet("Empleados", ["Empresa", "Cedula", "Nombre", "Puesto", "Fecha_Ingreso", "Fecha_Salida"])
    df_hist = cargar_datos_gsheet("Datos_Historicos", COLUMNAS_HISTORICO)

    empresas_list = sorted(df_emp["Empresa"].dropna().unique().tolist())

    if not empresas_list:
        st.warning("⚠️ No hay empresas registradas en Google Sheets. Vaya al Módulo 1 para agregarlas.")
    else:
        col_emp1, col_emp2 = st.columns(2)
        with col_emp1:
            empresa_activa = st.selectbox("Seleccione la Empresa:", empresas_list)

        df_colabs = df_empld[df_empld["Empresa"] == empresa_activa]
        lista_colabs = sorted(df_colabs["Nombre"].dropna().unique().tolist())

        if not lista_colabs:
            st.warning(f"La empresa '{empresa_activa}' no tiene empleados en el catálogo.")
        else:
            st.markdown("---")
            
            col_s1, col_s2, col_s3 = st.columns(3)
            with col_s1:
                colaborador_sel = st.selectbox("Seleccione el Colaborador:", lista_colabs)
            
            datos_colab = df_colabs[df_colabs["Nombre"] == colaborador_sel].iloc[-1].to_dict()
            
            with col_s2:
                st.text_input("Número de Cédula:", value=str(datos_colab.get("Cedula", "")), disabled=True)
            with col_s3:
                st.text_input("Puesto que Desempeña:", value=str(datos_colab.get("Puesto", "")), disabled=True)

            col_p1, col_p2, col_p3, col_p4 = st.columns(4)
            with col_p1:
                anio = st.number_input("Año:", min_value=2020, max_value=2035, value=2026)
            with col_p2:
                mes = st.selectbox("Mes:", MESES)
            with col_p3:
                moneda = st.selectbox("Moneda:", ["₡ (Colones)", "$ (Dólares)"])
            with col_p4:
                modo = st.selectbox("Periodicidad:", ["Mensual", "Quincenal", "Semanal"])

            fecha_consulta = datetime.date(int(anio), MESES_MAP[mes], 1)
            tc_bccr_auto = obtener_tipo_cambio_compra_bccr(fecha_consulta)
            tc_bccr = st.number_input(f"T.C. COMPRA BCCR ({mes} {anio}):", min_value=1.0, value=float(tc_bccr_auto), step=0.5, format="%.2f")

            st.subheader("4. Monto Ganado (Salario)")
            simbolo = "₡" if "Colones" in moneda else "$"
            
            monto_digitado = st.number_input(
                f"Digite el Salario Ganado ({simbolo}) [{modo}]:",
                min_value=0.0,
                step=1000.0,
                format="%.2f"
            )

            # CÁLCULOS TÉCNICOS AUTOMÁTICOS
            if modo == "Semanal":
                salario_base = monto_digitado * 4.3333
            elif modo == "Quincenal":
                salario_base = monto_digitado * 2.0
            else:
                salario_base = monto_digitado

            if "Dólares" in moneda:
                salario_usd = salario_base
                salario_crc = salario_base * tc_bccr
            else:
                salario_crc = salario_base
                salario_usd = salario_base / tc_bccr if tc_bccr > 0 else 0.0

            ccss_obrero_crc = salario_crc * PORCENTAJE_CCSS_OBRERO
            salario_neto_crc = salario_crc - ccss_obrero_crc
            ccss_patronal_crc = salario_crc * PORCENTAJE_CCSS_PATRONAL
            aguinaldo_crc = salario_crc * PORCENTAJE_AGUINALDO
            vacaciones_crc = salario_crc * PORCENTAJE_VACACIONES
            cesantia_crc = salario_crc * PORCENTAJE_CESANTIA
            total_prestaciones_crc = aguinaldo_crc + vacaciones_crc + cesantia_crc

            st.markdown(
                f"""
                <div style="background-color:#EBF5FB; padding:15px; border-radius:8px; border-left:5px solid #1F4E79; margin-top:10px;">
                    <h4 style="margin:0; color:#1F4E79;"><strong>Cálculo Inmediato de Planilla y Prestaciones:</strong></h4>
                    <p style="font-size:15px; margin:5px 0;">
                        <b>• Salario Mensual Bruto:</b> ₡{salario_crc:,.2f} CRC (${salario_usd:,.2f} USD)<br>
                        <b>• CCSS Obrero (10.67%):</b> ₡{ccss_obrero_crc:,.2f} CRC | <b>Salario Neto a Pagar:</b> ₡{salario_neto_crc:,.2f} CRC<br>
                        <b>• CCSS Carga Patronal (26.50%):</b> ₡{ccss_patronal_crc:,.2f} CRC
                    </p>
                    <hr style="margin:8px 0;">
                    <p style="font-size:15px; margin:0; color:#1B4F72;">
                        <b>• Aguinaldo (8.33%):</b> ₡{aguinaldo_crc:,.2f} | 
                        <b>• Vacaciones (4.16%):</b> ₡{vacaciones_crc:,.2f} | 
                        <b>• Cesantía (5.33%):</b> ₡{cesantia_crc:,.2f}<br>
                        <b>👉 Total Provisiones:</b> ₡{total_prestaciones_crc:,.2f} CRC
                    </p>
                </div>
                """,
                unsafe_allow_html=True
            )

            st.markdown("---")

            if st.button("💾 Guardar y Sincronizar Registro"):
                if monto_digitado <= 0:
                    st.error("El salario digitado debe ser mayor a 0.")
                else:
                    st.session_state["confirmar_guardado"] = {
                        "Empresa": empresa_activa,
                        "Año": int(anio),
                        "Mes": mes,
                        "Cedula": datos_colab.get("Cedula", ""),
                        "Nombre": colaborador_sel,
                        "Puesto": datos_colab.get("Puesto", ""),
                        "Fecha_Ingreso": datos_colab.get("Fecha_Ingreso", ""),
                        "Fecha_Salida": datos_colab.get("Fecha_Salida", "Activo"),
                        "Moneda": "USD" if "Dólares" in moneda else "CRC",
                        "Tipo_Cambio": tc_bccr,
                        "Salario_Ingresado": monto_digitado,
                        "Salario_Mensual_CRC": salario_crc,
                        "Salario_Mensual_USD": salario_usd,
                        "CCSS_Obrero_CRC": ccss_obrero_crc,
                        "Salario_Neto_CRC": salario_neto_crc,
                        "CCSS_Patronal_CRC": ccss_patronal_crc,
                        "Aguinaldo_CRC": aguinaldo_crc,
                        "Vacaciones_CRC": vacaciones_crc,
                        "Cesantia_CRC": cesantia_crc,
                        "Total_Prestaciones_CRC": total_prestaciones_crc
                    }

            if "confirmar_guardado" in st.session_state and st.session_state["confirmar_guardado"]:
                data_c = st.session_state["confirmar_guardado"]
                
                st.warning("⚠️ **VERIFICACIÓN DE DATOS ANTES DE GUARDAR EN LA NUBE**")
                st.write(f"• **Empresa:** {data_c['Empresa']} | **Empleado:** {data_c['Nombre']} ({data_c['Cedula']})")
                st.write(f"• **Periodo:** {data_c['Mes']} {data_c['Año']} | **Salario Mensual:** ₡{data_c['Salario_Mensual_CRC']:,.2f} CRC")

                col_conf1, col_conf2 = st.columns(2)
                with col_conf1:
                    if st.button("✅ Confirmar Guardado en Google Sheets"):
                        nuevo_reg = pd.DataFrame([data_c])
                        df_hist_updated = pd.concat([df_hist, nuevo_reg], ignore_index=True)
                        if guardar_datos_gsheet(df_hist_updated, "Datos_Historicos"):
                            st.session_state["confirmar_guardado"] = None
                            st.success(f"✅ ¡Registro guardado y sincronizado para **{data_c['Nombre']}**!")
                            st.rerun()

                with col_conf2:
                    if st.button("❌ Cancelar / Corregir"):
                        st.session_state["confirmar_guardado"] = None
                        st.rerun()

# --- MÓDULO 3: REPORTE DE PAGO DE CAJA POR EMPRESA ---
elif menu == "3. Costo CCSS Patronal por Empresa":
    st.title("3. Reporte de Pago a la Caja (CCSS) por Empresa")
    st.caption("Consulta sincronizada directamente de Google Sheets.")

    df_hist = cargar_datos_gsheet("Datos_Historicos", COLUMNAS_HISTORICO)
    df_hist = asegurar_calculos_historico(df_hist)

    if df_hist.empty:
        st.warning("No hay registros almacenados.")
    else:
        empresas_disponibles = sorted(df_hist["Empresa"].dropna().unique())
        empresa_sel = st.selectbox("Seleccione la Empresa:", empresas_disponibles)

        df_emp = df_hist[df_hist["Empresa"] == empresa_sel].copy()
        df_emp["Pago_Total_Caja_CRC"] = df_emp["CCSS_Obrero_CRC"] + df_emp["CCSS_Patronal_CRC"]
        df_emp["Pago_Total_Caja_USD"] = df_emp["Pago_Total_Caja_CRC"] / df_emp["Tipo_Cambio"]

        df_ccss_resumen = (
            df_emp.groupby(["Cedula", "Nombre", "Puesto"])
            .agg(
                Salario_CRC=("Salario_Mensual_CRC", "sum"),
                Salario_USD=("Salario_Mensual_USD", "sum"),
                Retencion_Obrero_CRC=("CCSS_Obrero_CRC", "sum"),
                Aporte_Patronal_CRC=("CCSS_Patronal_CRC", "sum"),
                Pago_Total_Caja_CRC=("Pago_Total_Caja_CRC", "sum"),
                Pago_Total_Caja_USD=("Pago_Total_Caja_USD", "sum"),
            )
            .reset_index()
        )

        st.subheader(f"Desglose de Pago de la CCSS - {empresa_sel}")
        st.dataframe(
            df_ccss_resumen.style.format(
                {
                    "Salario_CRC": "₡{:,.2f}",
                    "Salario_USD": "${:,.2f}",
                    "Retencion_Obrero_CRC": "₡{:,.2f}",
                    "Aporte_Patronal_CRC": "₡{:,.2f}",
                    "Pago_Total_Caja_CRC": "₡{:,.2f}",
                    "Pago_Total_Caja_USD": "${:,.2f}",
                }
            ),
            use_container_width=True,
        )

        st.subheader("✍️ Responsables de Firma para el Reporte")
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            elaborado_por = st.text_input("Elaborado por:", value="Oficial de Cumplimiento")
        with col_f2:
            revisado_por = st.text_input("Revisado por:", value="Gerencia General")

        if st.button("📄 Generar Reporte de CCSS en PDF"):
            data_pdf = generar_pdf_empresa(
                df_ccss_resumen,
                empresa_sel,
                "Reporte de Pago de Caja (CCSS) y Costos Patronales Bimoneda",
                elaborado_por,
                revisado_por,
            )
            st.download_button(
                label="🔴 Descargar Reporte CCSS (.pdf)",
                data=data_pdf,
                file_name=f"Reporte_CCSS_{empresa_sel}.pdf",
                mime="application/pdf",
            )

# --- MÓDULO 4: GESTOR DE REGISTROS HISTÓRICOS ---
elif menu == "4. Gestor de Registros Históricos":
    st.title("4. Modificar o Eliminar Registros Históricos")
    st.caption("Administre de forma directa la información en Google Sheets.")

    df_hist = cargar_datos_gsheet("Datos_Historicos", COLUMNAS_HISTORICO)
    df_hist = asegurar_calculos_historico(df_hist)

    if df_hist.empty:
        st.warning("No hay registros en la base de datos.")
    else:
        df_editado = st.data_editor(
            df_hist,
            num_rows="dynamic",
            use_container_width=True,
            column_config={
                "Mes": st.column_config.SelectboxColumn("Mes", options=MESES),
                "Moneda": st.column_config.SelectboxColumn("Moneda", options=["CRC", "USD"]),
                "Salario_Mensual_CRC": st.column_config.NumberColumn("Salario Mensual CRC", format="₡%.2f"),
                "Salario_Mensual_USD": st.column_config.NumberColumn("Salario Mensual USD", format="$%.2f"),
            },
        )

        if st.button("💾 Sincronizar Cambios en la Nube"):
            df_editado = asegurar_calculos_historico(df_editado)
            if guardar_datos_gsheet(df_editado, "Datos_Historicos"):
                st.success("✅ Base de datos histórica sincronizada en la nube.")
                st.rerun()

# --- MÓDULO 5: REPORTES Y EXPORTACIÓN ---
elif menu == "5. Reportes y Exportación":
    st.title("5. Reportes Acumulativos y Prestaciones")
    st.caption("Consolidado bimoneda sincronizado en tiempo real.")

    df_hist = cargar_datos_gsheet("Datos_Historicos", COLUMNAS_HISTORICO)
    df_hist = asegurar_calculos_historico(df_hist)

    if df_hist.empty:
        st.warning("No existen registros guardados en la base de datos.")
    else:
        df_hist["Cedula"] = df_hist["Cedula"].fillna("S/C").astype(str)
        df_hist["Nombre"] = df_hist["Nombre"].fillna("Sin Nombre").astype(str)
        df_hist["Puesto"] = df_hist["Puesto"].fillna("Sin Puesto").astype(str)
        df_hist["Fecha_Ingreso"] = df_hist["Fecha_Ingreso"].fillna("N/A").astype(str)
        df_hist["Fecha_Salida"] = df_hist["Fecha_Salida"].fillna("Activo").astype(str)

        empresas_disponibles = sorted(df_hist["Empresa"].dropna().unique())
        empresa_sel = st.selectbox("Seleccione la Empresa a Consultar:", empresas_disponibles)

        df_emp = df_hist[df_hist["Empresa"] == empresa_sel].copy()

        if df_emp.empty:
            st.warning(f"No hay registros salariales acumulados para la empresa '{empresa_sel}'.")
        else:
            df_emp["Total_Prestaciones_USD"] = df_emp["Total_Prestaciones_CRC"] / df_emp["Tipo_Cambio"]

            df_acum = (
                df_emp.groupby(
                    ["Cedula", "Nombre", "Puesto", "Fecha_Ingreso", "Fecha_Salida"], 
                    dropna=False
                )
                .agg(
                    Meses=("Mes", "count"),
                    Salario_CRC=("Salario_Mensual_CRC", "sum"),
                    Salario_USD=("Salario_Mensual_USD", "sum"),
                    CCSS_CRC=("CCSS_Obrero_CRC", "sum"),
                    Neto_CRC=("Salario_Neto_CRC", "sum"),
                    Aguinaldo_CRC=("Aguinaldo_CRC", "sum"),
                    Vacaciones_CRC=("Vacaciones_CRC", "sum"),
                    Cesantia_CRC=("Cesantia_CRC", "sum"),
                    Total_Prest_CRC=("Total_Prestaciones_CRC", "sum"),
                    Total_Prest_USD=("Total_Prestaciones_USD", "sum"),
                )
                .reset_index()
            )

            st.subheader(f"Resumen Acumulado de Planilla Bimoneda - {empresa_sel}")
            st.dataframe(
                df_acum.style.format(
                    {
                        "Salario_CRC": "₡{:,.2f}",
                        "Salario_USD": "${:,.2f}",
                        "CCSS_CRC": "₡{:,.2f}",
                        "Neto_CRC": "₡{:,.2f}",
                        "Aguinaldo_CRC": "₡{:,.2f}",
                        "Vacaciones_CRC": "₡{:,.2f}",
                        "Cesantia_CRC": "₡{:,.2f}",
                        "Total_Prest_CRC": "₡{:,.2f}",
                        "Total_Prest_USD": "${:,.2f}",
                    }
                ),
                use_container_width=True,
            )

            st.subheader("✍️ Responsables de Firma para los Reportes")
            col_f1, col_f2 = st.columns(2)
            with col_f1:
                elaborado_por = st.text_input("Elaborado por:", value="Oficial de Cumplimiento", key="el_rep")
            with col_f2:
                revisado_por = st.text_input("Revisado por:", value="Gerencia General", key="rev_rep")

            col_exp1, col_exp2, col_exp3 = st.columns(3)

            with col_exp1:
                data_excel = generar_excel_empresa(df_acum, empresa_sel, elaborado_por, revisado_por)
                st.download_button(
                    label="🟢 Descargar Excel (.xlsx)",
                    data=data_excel,
                    file_name=f"Reporte_{empresa_sel}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

            with col_exp2:
                data_word = generar_word_empresa(df_acum, empresa_sel, elaborado_por, revisado_por)
                st.download_button(
                    label="🔵 Descargar Word (.docx)",
                    data=data_word,
                    file_name=f"Reporte_{empresa_sel}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )

            with col_exp3:
                data_pdf = generar_pdf_empresa(
                    df_acum,
                    empresa_sel,
                    "Reporte Acumulado de Planilla, CCSS y Prestaciones Bimoneda",
                    elaborado_por,
                    revisado_por,
                )
                st.download_button(
                    label="🔴 Descargar PDF (.pdf)",
                    data=data_pdf,
                    file_name=f"Reporte_{empresa_sel}.pdf",
                    mime="application/pdf",
                )
