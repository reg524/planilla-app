import io
import re
import time
import xml.etree.ElementTree as ET
import pandas as pd
import requests
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
import streamlit as st
from streamlit_gsheets import GSheetsConnection

# Porcentajes según normativa de Costa Rica
PORCENTAJE_CCSS_OBRERO = 0.1067    # 10.67% CCSS Trabajador
PORCENTAJE_CCSS_PATRONAL = 0.2650  # 26.50% Cargas Sociales Patronales
PORCENTAJE_AGUINALDO = 1 / 12      # 8.33% Aguinaldo
PORCENTAJE_VACACIONES = 1 / 24     # 4.16% Vacaciones
PORCENTAJE_CESANTIA = 0.0533       # 5.33% Cesantía

MESES_MAP = {
    "Enero": 1, "Febrero": 2, "Marzo": 3, "Abril": 4, "Mayo": 5, "Junio": 6,
    "Julio": 7, "Agosto": 8, "Setiembre": 9, "Octubre": 10, "Noviembre": 11, "Diciembre": 12
}
MESES = list(MESES_MAP.keys())

COLUMNAS_HISTORICO = [
    "Empresa", "Año", "Mes", "Cedula", "Nombre", "Puesto", "Fecha_Ingreso", "Fecha_Salida",
    "Detalle", "Moneda", "Tipo_Cambio", "Salario_Ingresado", "Salario_Mensual_CRC", "Salario_Mensual_USD",
    "CCSS_Obrero_CRC", "Salario_Neto_CRC", "CCSS_Patronal_CRC",
    "Aguinaldo_CRC", "Vacaciones_CRC", "Cesantia_CRC", "Total_Prestaciones_CRC"
]

@st.cache_resource
def get_gsheet_connection():
    return st.connection("gsheets", type=GSheetsConnection)

def limpiar_nombre_archivo(texto):
    return re.sub(r'[\\/*?:"<>|]', "_", str(texto)).strip()

def limpiar_numero_serie(serie):
    return (
        serie.astype(str)
        .str.replace("₡", "", regex=False)
        .str.replace("$", "", regex=False)
        .str.replace(",", "", regex=False)
        .str.replace(" ", "", regex=False)
        .str.strip()
        .pipe(pd.to_numeric, errors="coerce")
        .fillna(0.0)
    )

def cargar_datos_gsheet(worksheet_name, columnas_defecto):
    cache_key = f"backup_df_{worksheet_name}"
    conn = get_gsheet_connection()
    
    for _ in range(2):
        try:
            df = conn.read(worksheet=worksheet_name, ttl=60)
            df = df.dropna(how="all")
            for col in columnas_defecto:
                if col not in df.columns:
                    df[col] = 0.0 if ("CRC" in col or "USD" in col or col in ["Tipo_Cambio", "Salario_Ingresado"]) else ""
            
            cols_texto = ["Empresa", "Cedula", "Cedula_Juridica", "Nombre", "Puesto", "Contacto", "Fecha_Ingreso", "Fecha_Salida", "Detalle", "Mes", "Moneda"]
            for col in cols_texto:
                if col in df.columns:
                    df[col] = df[col].astype(str).replace(["nan", "None", "<NA>", "NaN"], "").str.strip()
            
            df = df.reset_index(drop=True)
            st.session_state[cache_key] = df.copy()
            return df
        except Exception:
            time.sleep(0.5)
    
    if cache_key in st.session_state and not st.session_state[cache_key].empty:
        return st.session_state[cache_key].copy()
        
    return pd.DataFrame(columns=columnas_defecto)

def guardar_datos_gsheet(df, worksheet_name):
    try:
        conn = get_gsheet_connection()
        df_clean = df.copy().dropna(how="all")
        for col in df_clean.columns:
            if df_clean[col].dtype == object:
                df_clean[col] = df_clean[col].astype(str).replace(["nan", "None", "<NA>", "NaN"], "").fillna("")
        conn.update(worksheet=worksheet_name, data=df_clean)
        
        st.session_state[f"backup_df_{worksheet_name}"] = df_clean.reset_index(drop=True).copy()
        st.cache_data.clear()
        return True
    except Exception as e:
        st.error(f"Error al guardar en Google Sheets: {e}")
        return False

def normalizar_mes_nombre(mes_val):
    m = str(mes_val).strip().capitalize()
    if m in ["Septiembre", "9", "09"]:
        return "Setiembre"
    num_to_mes = {str(v): k for k, v in MESES_MAP.items()}
    num_to_mes_pad = {f"{v:02d}": k for k, v in MESES_MAP.items()}
    if m in num_to_mes:
        return num_to_mes[m]
    if m in num_to_mes_pad:
        return num_to_mes_pad[m]
    return m if m in MESES_MAP else "Enero"

def asegurar_calculos_historico(df):
    if df.empty:
        return df

    if "Detalle" not in df.columns:
        df["Detalle"] = ""

    # Limpieza segura del Año (maneja "2024.0", "2024", espacios, etc.)
    anio_limpio = df.get("Año", "2026").astype(str).str.replace(".0", "", regex=False).str.strip()
    df["Año_Num"] = pd.to_numeric(anio_limpio, errors="coerce").fillna(2026).astype(int)
    df["Año"] = df["Año_Num"]

    df["Mes"] = df.get("Mes", "Enero").apply(normalizar_mes_nombre)
    df["Mes_Num"] = df["Mes"].map(MESES_MAP).fillna(1).astype(int)
    df["Periodo_Orden"] = (df["Año_Num"] * 100) + df["Mes_Num"]

    df["Salario_Mensual_CRC"] = limpiar_numero_serie(df.get("Salario_Mensual_CRC", pd.Series(0.0, index=df.index)))
    df["Salario_Mensual_USD"] = limpiar_numero_serie(df.get("Salario_Mensual_USD", pd.Series(0.0, index=df.index)))
    df["Salario_Ingresado"] = limpiar_numero_serie(df.get("Salario_Ingresado", pd.Series(0.0, index=df.index)))
    df["Tipo_Cambio"] = limpiar_numero_serie(df.get("Tipo_Cambio", pd.Series(515.0, index=df.index)))
    df["Tipo_Cambio"] = df["Tipo_Cambio"].apply(lambda x: x if x > 0 else 515.0)

    # Reconstruir Salario_Mensual_CRC si registros antiguos lo traen en 0
    mask_crc_cero = df["Salario_Mensual_CRC"] <= 0
    if mask_crc_cero.any():
        es_usd = df["Moneda"].astype(str).str.upper().str.contains("USD|DOLAR")
        df.loc[mask_crc_cero & es_usd & (df["Salario_Mensual_USD"] > 0), "Salario_Mensual_CRC"] = (
            df["Salario_Mensual_USD"] * df["Tipo_Cambio"]
        )
        mask_crc_cero = df["Salario_Mensual_CRC"] <= 0
        df.loc[mask_crc_cero & es_usd & (df["Salario_Ingresado"] > 0), "Salario_Mensual_CRC"] = (
            df["Salario_Ingresado"] * df["Tipo_Cambio"]
        )
        mask_crc_cero = df["Salario_Mensual_CRC"] <= 0
        df.loc[mask_crc_cero & (~es_usd) & (df["Salario_Ingresado"] > 0), "Salario_Mensual_CRC"] = (
            df["Salario_Ingresado"]
        )

    mask_usd_cero = (df["Salario_Mensual_USD"] <= 0) & (df["Salario_Mensual_CRC"] > 0)
    df.loc[mask_usd_cero, "Salario_Mensual_USD"] = df.loc[mask_usd_cero, "Salario_Mensual_CRC"] / df.loc[mask_usd_cero, "Tipo_Cambio"]

    df["CCSS_Obrero_CRC"] = df["Salario_Mensual_CRC"] * PORCENTAJE_CCSS_OBRERO
    df["Salario_Neto_CRC"] = df["Salario_Mensual_CRC"] - df["CCSS_Obrero_CRC"]
    df["CCSS_Patronal_CRC"] = df["Salario_Mensual_CRC"] * PORCENTAJE_CCSS_PATRONAL
    df["Aguinaldo_CRC"] = df["Salario_Mensual_CRC"] * PORCENTAJE_AGUINALDO
    df["Vacaciones_CRC"] = df["Salario_Mensual_CRC"] * PORCENTAJE_VACACIONES
    df["Cesantia_CRC"] = df["Salario_Mensual_CRC"] * PORCENTAJE_CESANTIA
    df["Total_Prestaciones_CRC"] = df["Aguinaldo_CRC"] + df["Vacaciones_CRC"] + df["Cesantia_CRC"]

    return df

@st.cache_data(ttl=3600)
def obtener_tipo_cambio_compra_bccr(fecha_dt):
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

def generar_excel_empresa(df_emp, nombre_empresa):
    output = io.BytesIO()
    hoja = limpiar_nombre_archivo(nombre_empresa)[:30] or "Reporte"
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_emp.to_excel(writer, sheet_name=hoja, index=False)
    return output.getvalue()

def generar_word_empresa(df_emp, nombre_empresa, titulo_reporte, elaborado_por, revisado_por):
    doc = Document()
    p_title = doc.add_heading(f"{nombre_empresa}\n{titulo_reporte}", 0)
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    table = doc.add_table(rows=1, cols=len(df_emp.columns))
    table.style = "Table Grid"

    hdr_cells = table.rows[0].cells
    for i, col_name in enumerate(df_emp.columns):
        hdr_cells[i].text = str(col_name).replace("_", " ")

    for _, row in df_emp.iterrows():
        row_cells = table.add_row().cells
        for i, (col_name, val) in enumerate(zip(df_emp.columns, row)):
            if isinstance(val, float):
                row_cells[i].text = f"${val:,.2f}" if "USD" in str(col_name) else f"₡{val:,.2f}"
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
        fontSize=12,
        alignment=1,
        spaceAfter=6
    )
    style_cell = ParagraphStyle('CellText', fontName='Helvetica', fontSize=6.5, leading=8, alignment=1)
    style_header = ParagraphStyle('HeaderStyle', fontName='Helvetica-Bold', fontSize=6.5, leading=8, textColor=colors.whitesmoke, alignment=1)

    elements.append(Paragraph(f"<b>{str(nombre_empresa).upper()}</b>", style_title))
    elements.append(Paragraph(f"<b>{titulo_reporte}</b>", style_title))
    elements.append(Spacer(1, 8))

    headers = [Paragraph(str(col).replace("_", " "), style_header) for col in df_emp.columns]
    data = [headers]

    for _, row in df_emp.iterrows():
        fila_formateada = []
        for col_name, val in zip(df_emp.columns, row):
            if isinstance(val, float):
                texto = f"${val:,.2f}" if "USD" in str(col_name) else f"₡{val:,.2f}"
            else:
                texto = str(val) if pd.notnull(val) else ""
            fila_formateada.append(Paragraph(texto, style_cell))
        data.append(fila_formateada)

    num_cols = max(len(df_emp.columns), 1)
    col_width = 762.0 / num_cols

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
