import datetime
import pandas as pd
import streamlit as st
from reportes_planilla import (
    COLUMNAS_EMPLEADOS,
    COLUMNAS_EMPRESAS,
    COLUMNAS_HISTORICO,
    MESES,
    MESES_MAP,
    PORCENTAJE_AGUINALDO,
    PORCENTAJE_CCSS_OBRERO,
    PORCENTAJE_CCSS_PATRONAL,
    PORCENTAJE_CESANTIA,
    PORCENTAJE_VACACIONES,
    asegurar_calculos_historico,
    cargar_datos_gsheet,
    generar_backup_completo_excel,
    generar_excel_empresa,
    generar_pdf_empresa,
    generar_word_empresa,
    guardar_copia_local_opcional,
    guardar_datos_gsheet,
    limpiar_nombre_archivo,
    obtener_tipo_cambio_compra_bccr,
    restaurar_backup_completo_excel,
)

st.set_page_config(
    page_title="Sistema de Planilla Bimoneda, Multiempresa y Cargas Sociales",
    page_icon="⚖️",
    layout="wide",
)

def limpiar_columnas_agrupacion(df, mapeo_defecto):
    """Garantiza que no existan valores None/NaN en columnas de agrupación de Pandas"""
    df_out = df.copy()
    for col, valor_defecto in mapeo_defecto.items():
        if col not in df_out.columns:
            df_out[col] = valor_defecto
        else:
            df_out[col] = (
                df_out[col]
                .fillna(valor_defecto)
                .astype(str)
                .str.strip()
                .replace(["", "None", "nan", "NaN", "<NA>", "null"], valor_defecto)
            )
    return df_out

# --- BARRA LATERAL ---
st.sidebar.title("Módulos del Sistema (Sincronizado)")

if st.sidebar.button("🔄 Actualizar Datos de la Nube", use_container_width=True):
    st.cache_data.clear()
    for k in list(st.session_state.keys()):
        if k.startswith("backup_df_"):
            del st.session_state[k]
    st.rerun()

menu = st.sidebar.radio(
    "Seleccione el Menú:",
    [
        "1. Catálogo de Empresas y Empleados",
        "2. Salarios a Ingresar (Planilla)",
        "3. Costo CCSS Patronal por Empresa",
        "4. Gestor de Registros Históricos",
        "5. Reportes y Exportación",
        "6. Respaldo y Restauración (Backup)",
    ],
)

# --- MÓDULO 1: CATÁLOGO DE EMPRESAS Y EMPLEADOS ---
if menu == "1. Catálogo de Empresas y Empleados":
    st.title("1. Administración de Empresas y Colaboradores")
    st.caption("Presione Enter o haga clic fuera de la celda antes de presionar 'Sincronizar'.")

    df_emp = cargar_datos_gsheet("Empresas", COLUMNAS_EMPRESAS)
    df_empld = cargar_datos_gsheet("Empleados", COLUMNAS_EMPLEADOS)

    tab1, tab2 = st.tabs(["🏢 Catálogo de Empresas", "👤 Padrón de Empleados por Empresa"])

    with tab1:
        st.subheader("Base de Datos de Empresas (Nube)")
        df_emp_edited = st.data_editor(
            df_emp,
            num_rows="dynamic",
            use_container_width=True,
            key="editor_empresas",
            column_config={
                "Empresa": st.column_config.TextColumn("Nombre de la Empresa"),
                "Cedula_Juridica": st.column_config.TextColumn("Cédula Jurídica / Dimex"),
                "Contacto": st.column_config.TextColumn("Contacto / Observaciones"),
            }
        )
        if st.button("💾 Sincronizar Catálogo de Empresas"):
            df_emp_clean = df_emp_edited[df_emp_edited["Empresa"].astype(str).str.strip() != ""]
            if guardar_datos_gsheet(df_emp_clean, "Empresas"):
                st.success(f"✅ Catálogo sincronizado ({len(df_emp_clean)} empresas guardadas).")
                st.rerun()

    with tab2:
        st.subheader("Base de Datos de Empleados (Nube)")
        empresas_list = sorted([str(e) for e in df_emp["Empresa"].dropna().unique() if str(e).strip() != ""])
        
        if not empresas_list:
            st.warning("Primero registre al menos una empresa en el catálogo de la pestaña anterior.")
        else:
            emp_sel_padron = st.selectbox("Seleccione la Empresa:", empresas_list)
            df_sub_empld = df_empld[df_empld["Empresa"] == emp_sel_padron].copy().reset_index(drop=True)
            
            columnas_padron = ["Cedula", "Nombre", "Puesto", "Fecha_Ingreso", "Fecha_Salida"]
            for c in columnas_padron:
                if c not in df_sub_empld.columns:
                    df_sub_empld[c] = ""
            df_sub_empld = df_sub_empld[columnas_padron]

            df_sub_edited = st.data_editor(
                df_sub_empld,
                num_rows="dynamic",
                use_container_width=True,
                key=f"editor_empld_{emp_sel_padron}",
                column_config={
                    "Cedula": st.column_config.TextColumn("Cédula Colaborador"),
                    "Nombre": st.column_config.TextColumn("Nombre Completo"),
                    "Puesto": st.column_config.TextColumn("Puesto / Cargo"),
                    "Fecha_Ingreso": st.column_config.TextColumn("Fecha Ingreso (DD/MM/AAAA)"),
                    "Fecha_Salida": st.column_config.TextColumn("Fecha Salida / Activo"),
                }
            )

            if st.button(f"💾 Sincronizar Padrón de {emp_sel_padron}"):
                df_sub_clean = df_sub_edited[
                    (df_sub_edited["Nombre"].astype(str).str.strip() != "") |
                    (df_sub_edited["Cedula"].astype(str).str.strip() != "")
                ].copy()
                
                df_otros = df_empld[df_empld["Empresa"] != emp_sel_padron].copy()
                df_sub_clean["Empresa"] = emp_sel_padron
                
                df_final_empld = pd.concat([df_otros, df_sub_clean], ignore_index=True)[COLUMNAS_EMPLEADOS]
                
                if guardar_datos_gsheet(df_final_empld, "Empleados"):
                    st.success(f"✅ Personal de **{emp_sel_padron}** sincronizado ({len(df_sub_clean)} colaboradores guardados).")
                    st.rerun()

# --- MÓDULO 2: SALARIOS A INGRESAR (PLANILLA) ---
elif menu == "2. Salarios a Ingresar (Planilla)":
    st.title("2. Salarios a Ingresar")
    st.caption("Alimentación de salarios con sincronización remota instantánea.")

    df_emp = cargar_datos_gsheet("Empresas", COLUMNAS_EMPRESAS)
    df_empld = cargar_datos_gsheet("Empleados", COLUMNAS_EMPLEADOS)

    empresas_list = sorted([str(e) for e in df_emp["Empresa"].dropna().unique() if str(e).strip() != ""])

    if not empresas_list:
        st.warning("⚠️ No se detectaron empresas cargadas desde Google Sheets.")
        if st.button("🔄 Reintentar conexión con Google Sheets"):
            st.cache_data.clear()
            st.rerun()
    else:
        col_emp1, col_emp2 = st.columns(2)
        with col_emp1:
            empresa_activa = st.selectbox("Seleccione la Empresa:", empresas_list, key="sel_emp_mod2")

        df_colabs = df_empld[df_empld["Empresa"] == empresa_activa]
        lista_colabs = sorted([str(n) for n in df_colabs["Nombre"].dropna().unique() if str(n).strip() != ""])

        if not lista_colabs:
            st.warning(f"La empresa '{empresa_activa}' no tiene empleados en el catálogo.")
        else:
            st.markdown("---")
            
            col_s1, col_s2, col_s3 = st.columns(3)
            with col_s1:
                colaborador_sel = st.selectbox("Seleccione el Colaborador:", lista_colabs, key="sel_colab_mod2")
            
            datos_colab = df_colabs[df_colabs["Nombre"] == colaborador_sel].iloc[-1].to_dict()
            
            with col_s2:
                st.text_input("Número de Cédula:", value=str(datos_colab.get("Cedula", "")), disabled=True)
            with col_s3:
                st.text_input("Puesto que Desempeña:", value=str(datos_colab.get("Puesto", "")), disabled=True)

            col_p1, col_p2, col_p3, col_p4 = st.columns(4)
            with col_p1:
                anio = st.number_input("Año:", min_value=2015, max_value=2035, value=2026)
            with col_p2:
                mes = st.selectbox("Mes:", MESES)
            with col_p3:
                moneda = st.selectbox("Moneda:", ["₡ (Colones)", "$ (Dólares)"])
            with col_p4:
                modo = st.selectbox("Periodicidad:", ["Mensual", "Quincenal", "Semanal"])

            col_tc, col_det = st.columns([1, 2])
            with col_tc:
                fecha_consulta = datetime.date(int(anio), MESES_MAP[mes], 1)
                tc_bccr_auto = obtener_tipo_cambio_compra_bccr(fecha_consulta)
                tc_bccr = st.number_input(f"T.C. COMPRA BCCR ({mes} {anio}):", min_value=1.0, value=float(tc_bccr_auto), step=0.5, format="%.2f")
            with col_det:
                detalle_ingreso = st.text_input(
                    "Detalle / Concepto del Pago:",
                    placeholder="Ej: Salario ordinario, Horas extra, Bonificación, Liquidación..."
                )

            st.subheader("4. Monto Ganado (Salario)")
            simbolo = "₡" if "Colones" in moneda else "$"
            
            monto_digitado = st.number_input(
                f"Digite el Salario Ganado ({simbolo}) [{modo}]:",
                min_value=0.0,
                step=1000.0,
                format="%.2f"
            )

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
                        "Detalle": detalle_ingreso.strip() if detalle_ingreso.strip() else "Salario Ordinario",
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
                st.write(f"• **Periodo:** {data_c['Mes']} {data_c['Año']} | **Detalle:** {data_c['Detalle']}")
                st.write(f"• **Salario Mensual:** ₡{data_c['Salario_Mensual_CRC']:,.2f} CRC")

                col_conf1, col_conf2 = st.columns(2)
                with col_conf1:
                    if st.button("✅ Confirmar Guardado en Google Sheets"):
                        df_hist = cargar_datos_gsheet("Datos_Historicos", COLUMNAS_HISTORICO)
                        nuevo_reg = pd.DataFrame([data_c])
                        df_hist_updated = pd.concat([df_hist, nuevo_reg], ignore_index=True)
                        for col in COLUMNAS_HISTORICO:
                            if col not in df_hist_updated.columns:
                                df_hist_updated[col] = ""
                        df_hist_updated = df_hist_updated[COLUMNAS_HISTORICO]

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
    st.caption("Cálculo automático por rango de fechas mensuales desde los registros históricos.")

    df_hist = cargar_datos_gsheet("Datos_Historicos", COLUMNAS_HISTORICO)
    df_hist = asegurar_calculos_historico(df_hist)

    if df_hist.empty:
        st.error("❌ La hoja 'Datos_Historicos' en Google Sheets se leyó vacía (0 filas).")
    else:
        df_hist = limpiar_columnas_agrupacion(
            df_hist,
            {
                "Empresa": "Sin Empresa",
                "Cedula": "S/C",
                "Nombre": "Sin Nombre",
                "Puesto": "Sin Puesto",
            },
        )

        empresas_disponibles = sorted([e for e in df_hist["Empresa"].unique() if e != "Sin Empresa"])
        
        if not empresas_disponibles:
            st.warning(f"⚠️ Se leyeron {len(df_hist)} filas en 'Datos_Historicos', pero la columna 'Empresa' está vacía.")
        else:
            empresa_sel = st.selectbox("Seleccione la Empresa:", empresas_disponibles, key="ccss_emp_sel")
            df_emp_base = df_hist[df_hist["Empresa"].str.lower() == empresa_sel.lower()].copy()

            st.markdown("#### 📅 Rango de Fechas Mensuales a Calcular (Puede digitar años anteriores)")
            col_r1, col_r2, col_r3, col_r4 = st.columns(4)
            with col_r1:
                mes_inicio = st.selectbox("Desde Mes:", MESES, index=0, key="ccss_mes_ini")
            with col_r2:
                anio_inicio = st.number_input("Desde Año:", min_value=2015, max_value=2035, value=2024, step=1, key="ccss_anio_ini")
            with col_r3:
                mes_fin = st.selectbox("Hasta Mes:", MESES, index=11, key="ccss_mes_fin")
            with col_r4:
                anio_fin = st.number_input("Hasta Año:", min_value=2015, max_value=2035, value=2026, step=1, key="ccss_anio_fin")

            periodo_ini = (int(anio_inicio) * 100) + MESES_MAP[mes_inicio]
            periodo_fin = (int(anio_fin) * 100) + MESES_MAP[mes_fin]

            if periodo_ini > periodo_fin:
                st.error("⚠️ El periodo inicial ('Desde') no puede ser posterior al periodo final ('Hasta').")
            else:
                df_emp = df_emp_base[
                    (df_emp_base["Periodo_Orden"] >= periodo_ini) &
                    (df_emp_base["Periodo_Orden"] <= periodo_fin)
                ].copy()

                if df_emp.empty:
                    anios_ex = sorted(df_emp_base["Año_Num"].unique().tolist())
                    st.info(f"No hay registros para **{empresa_sel}** en el rango seleccionado. Años con datos disponibles para esta empresa: **{anios_ex}**")
                else:
                    df_emp["Pago_Total_Caja_CRC"] = df_emp["CCSS_Obrero_CRC"] + df_emp["CCSS_Patronal_CRC"]
                    df_emp["Pago_Total_Caja_USD"] = df_emp["Pago_Total_Caja_CRC"] / df_emp["Tipo_Cambio"]

                    df_ccss_resumen = (
                        df_emp.groupby(["Cedula", "Nombre", "Puesto"], as_index=False, dropna=False)
                        .agg(
                            Meses=("Mes", "count"),
                            Salario_CRC=("Salario_Mensual_CRC", "sum"),
                            Salario_USD=("Salario_Mensual_USD", "sum"),
                            Retencion_Obrero_CRC=("CCSS_Obrero_CRC", "sum"),
                            Aporte_Patronal_CRC=("CCSS_Patronal_CRC", "sum"),
                            Pago_Total_Caja_CRC=("Pago_Total_Caja_CRC", "sum"),
                            Pago_Total_Caja_USD=("Pago_Total_Caja_USD", "sum"),
                        )
                    )

                    total_sal_crc = df_ccss_resumen["Salario_CRC"].sum()
                    total_obr_crc = df_ccss_resumen["Retencion_Obrero_CRC"].sum()
                    total_pat_crc = df_ccss_resumen["Aporte_Patronal_CRC"].sum()
                    total_caja_crc = df_ccss_resumen["Pago_Total_Caja_CRC"].sum()

                    st.markdown("---")
                    m1, m2, m3, m4 = st.columns(4)
                    m1.metric("Total Salarios del Rango", f"₡{total_sal_crc:,.2f}")
                    m2.metric("CCSS Obrero (10.67%)", f"₡{total_obr_crc:,.2f}")
                    m3.metric("CCSS Patronal (26.50%)", f"₡{total_pat_crc:,.2f}")
                    m4.metric("Total a Pagar a la Caja", f"₡{total_caja_crc:,.2f}")

                    st.subheader(f"Desglose de Pago de la CCSS ({mes_inicio} {int(anio_inicio)} - {mes_fin} {int(anio_fin)})")
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

                    with st.expander("🔍 Ver detalle mensual incluido en este cálculo"):
                        cols_detalle_ver = ["Año_Num", "Mes", "Cedula", "Nombre", "Detalle", "Salario_Mensual_CRC", "CCSS_Obrero_CRC", "CCSS_Patronal_CRC"]
                        st.dataframe(
                            df_emp[cols_detalle_ver].rename(columns={"Año_Num": "Año"}).style.format(
                                {
                                    "Salario_Mensual_CRC": "₡{:,.2f}",
                                    "CCSS_Obrero_CRC": "₡{:,.2f}",
                                    "CCSS_Patronal_CRC": "₡{:,.2f}",
                                }
                            ),
                            use_container_width=True,
                        )

                    st.subheader("✍️ Responsables de Firma y Descarga de Archivos")
                    col_f1, col_f2 = st.columns(2)
                    with col_f1:
                        elaborado_por = st.text_input("Elaborado por:", value="Lic. Roy F. Esquivel González", key="el_ccss")
                    with col_f2:
                        revisado_por = st.text_input("Revisado por:", value="Gerencia General", key="rev_ccss")

                    nombre_limpio_emp = limpiar_nombre_archivo(empresa_sel)
                    subtitulo_ccss = f"Reporte de Pago de Caja (CCSS) | Rango: {mes_inicio} {int(anio_inicio)} a {mes_fin} {int(anio_fin)}"
                    nom_pdf = f"Reporte_CCSS_{nombre_limpio_emp}_{mes_inicio}{int(anio_inicio)}_a_{mes_fin}{int(anio_fin)}.pdf"
                    nom_xlsx = f"Reporte_CCSS_{nombre_limpio_emp}_{mes_inicio}{int(anio_inicio)}_a_{mes_fin}{int(anio_fin)}.xlsx"
                    nom_docx = f"Reporte_CCSS_{nombre_limpio_emp}_{mes_inicio}{int(anio_inicio)}_a_{mes_fin}{int(anio_fin)}.docx"

                    try:
                        data_pdf_ccss = generar_pdf_empresa(df_ccss_resumen, empresa_sel, subtitulo_ccss, elaborado_por, revisado_por)
                        data_excel_ccss = generar_excel_empresa(df_ccss_resumen, empresa_sel)
                        data_word_ccss = generar_word_empresa(df_ccss_resumen, empresa_sel, subtitulo_ccss, elaborado_por, revisado_por)

                        ruta_local = guardar_copia_local_opcional(data_pdf_ccss, nom_pdf)
                        guardar_copia_local_opcional(data_excel_ccss, nom_xlsx)
                        guardar_copia_local_opcional(data_word_ccss, nom_docx)

                        if ruta_local:
                            st.success(f"📂 Copia local guardada automáticamente en: `{ruta_local}`")

                        col_btn1, col_btn2, col_btn3 = st.columns(3)
                        with col_btn1:
                            st.download_button(
                                label="🔴 Descargar Reporte CCSS (.pdf)",
                                data=data_pdf_ccss,
                                file_name=nom_pdf,
                                mime="application/pdf",
                                key="dl_pdf_ccss"
                            )
                        with col_btn2:
                            st.download_button(
                                label="🟢 Descargar Reporte CCSS (.xlsx)",
                                data=data_excel_ccss,
                                file_name=nom_xlsx,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key="dl_xlsx_ccss"
                            )
                        with col_btn3:
                            st.download_button(
                                label="🔵 Descargar Reporte CCSS (.docx)",
                                data=data_word_ccss,
                                file_name=nom_docx,
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                key="dl_docx_ccss"
                            )
                    except Exception as e:
                        st.error(f"Error técnico al construir el archivo del reporte: {e}")

# --- MÓDULO 4: GESTOR DE REGISTROS HISTÓRICOS ---
elif menu == "4. Gestor de Registros Históricos":
    st.title("4. Modificar o Eliminar Registros Históricos")
    st.caption("Administre de forma directa la información en Google Sheets.")

    df_hist = cargar_datos_gsheet("Datos_Historicos", COLUMNAS_HISTORICO)
    df_hist = asegurar_calculos_historico(df_hist)

    if df_hist.empty:
        st.warning("No hay registros en la base de datos.")
    else:
        df_mostrar = df_hist[COLUMNAS_HISTORICO].copy()
        df_editado = st.data_editor(
            df_mostrar,
            num_rows="dynamic",
            use_container_width=True,
            column_config={
                "Empresa": st.column_config.TextColumn("Empresa"),
                "Nombre": st.column_config.TextColumn("Nombre"),
                "Detalle": st.column_config.TextColumn("Detalle"),
                "Mes": st.column_config.SelectboxColumn("Mes", options=MESES),
                "Moneda": st.column_config.SelectboxColumn("Moneda", options=["CRC", "USD"]),
                "Salario_Mensual_CRC": st.column_config.NumberColumn("Salario Mensual CRC", format="₡%.2f"),
                "Salario_Mensual_USD": st.column_config.NumberColumn("Salario Mensual USD", format="$%.2f"),
            },
        )

        if st.button("💾 Sincronizar Cambios en la Nube"):
            df_editado = asegurar_calculos_historico(df_editado)
            df_guardar = df_editado[COLUMNAS_HISTORICO].copy()
            if guardar_datos_gsheet(df_guardar, "Datos_Historicos"):
                st.success("✅ Base de datos histórica sincronizada y recalculada en la nube.")
                st.rerun()

# --- MÓDULO 5: REPORTES Y EXPORTACIÓN ---
elif menu == "5. Reportes y Exportación":
    st.title("5. Reportes Acumulativos y Prestaciones")
    st.caption("Consolidado bimoneda por rango de fechas mensuales.")

    df_hist = cargar_datos_gsheet("Datos_Historicos", COLUMNAS_HISTORICO)
    df_hist = asegurar_calculos_historico(df_hist)

    if df_hist.empty:
        st.error("❌ La hoja 'Datos_Historicos' en Google Sheets se leyó vacía (0 filas).")
    else:
        df_hist = limpiar_columnas_agrupacion(
            df_hist,
            {
                "Empresa": "Sin Empresa",
                "Cedula": "S/C",
                "Nombre": "Sin Nombre",
                "Puesto": "Sin Puesto",
                "Fecha_Ingreso": "N/A",
                "Fecha_Salida": "Activo",
            },
        )

        empresas_disponibles = sorted([e for e in df_hist["Empresa"].unique() if e != "Sin Empresa"])
        if not empresas_disponibles:
            st.warning(f"⚠️ Se leyeron {len(df_hist)} filas, pero no hay nombres de empresa en la columna 'Empresa'.")
        else:
            empresa_sel = st.selectbox("Seleccione la Empresa a Consultar:", empresas_disponibles, key="rep_emp_sel")

            df_emp_base = df_hist[df_hist["Empresa"].str.lower() == empresa_sel.lower()].copy()

            st.markdown("#### 📅 Rango de Fechas Mensuales a Consultar (Puede digitar años anteriores)")
            col_r1, col_r2, col_r3, col_r4 = st.columns(4)
            with col_r1:
                mes_inicio = st.selectbox("Desde Mes:", MESES, index=0, key="rep_mes_ini")
            with col_r2:
                anio_inicio = st.number_input("Desde Año:", min_value=2015, max_value=2035, value=2024, step=1, key="rep_anio_ini")
            with col_r3:
                mes_fin = st.selectbox("Hasta Mes:", MESES, index=11, key="rep_mes_fin")
            with col_r4:
                anio_fin = st.number_input("Hasta Año:", min_value=2015, max_value=2035, value=2026, step=1, key="rep_anio_fin")

            periodo_ini = (int(anio_inicio) * 100) + MESES_MAP[mes_inicio]
            periodo_fin = (int(anio_fin) * 100) + MESES_MAP[mes_fin]

            if periodo_ini > periodo_fin:
                st.error("⚠️ El periodo inicial ('Desde') no puede ser posterior al periodo final ('Hasta').")
            else:
                df_emp = df_emp_base[
                    (df_emp_base["Periodo_Orden"] >= periodo_ini) &
                    (df_emp_base["Periodo_Orden"] <= periodo_fin)
                ].copy()

                if df_emp.empty:
                    anios_ex = sorted(df_emp_base["Año_Num"].unique().tolist())
                    st.warning(f"No hay registros salariales para **{empresa_sel}** en ese rango. Años con datos disponibles: **{anios_ex}**")
                else:
                    df_emp["Total_Prestaciones_USD"] = df_emp["Total_Prestaciones_CRC"] / df_emp["Tipo_Cambio"]

                    df_acum = (
                        df_emp.groupby(
                            ["Cedula", "Nombre", "Puesto", "Fecha_Ingreso", "Fecha_Salida"], 
                            as_index=False,
                            dropna=False,
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
                    )

                    st.subheader(f"Resumen Acumulado ({mes_inicio} {int(anio_inicio)} - {mes_fin} {int(anio_fin)}) - {empresa_sel}")
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

                    st.subheader("✍️ Responsables de Firma y Descarga de Archivos")
                    col_f1, col_f2 = st.columns(2)
                    with col_f1:
                        elaborado_por = st.text_input("Elaborado por:", value="Lic. Roy F. Esquivel González", key="el_rep")
                    with col_f2:
                        revisado_por = st.text_input("Revisado por:", value="Gerencia General", key="rev_rep")

                    nombre_limpio_emp = limpiar_nombre_archivo(empresa_sel)
                    subtitulo_rep = f"Reporte Acumulado de Planilla, CCSS y Prestaciones ({mes_inicio} {int(anio_inicio)} - {mes_fin} {int(anio_fin)})"
                    nom_xlsx = f"Reporte_{nombre_limpio_emp}_{mes_inicio}{int(anio_inicio)}_a_{mes_fin}{int(anio_fin)}.xlsx"
                    nom_docx = f"Reporte_{nombre_limpio_emp}_{mes_inicio}{int(anio_inicio)}_a_{mes_fin}{int(anio_fin)}.docx"
                    nom_pdf = f"Reporte_{nombre_limpio_emp}_{mes_inicio}{int(anio_inicio)}_a_{mes_fin}{int(anio_fin)}.pdf"

                    try:
                        data_excel = generar_excel_empresa(df_acum, empresa_sel)
                        data_word = generar_word_empresa(df_acum, empresa_sel, subtitulo_rep, elaborado_por, revisado_por)
                        data_pdf = generar_pdf_empresa(df_acum, empresa_sel, subtitulo_rep, elaborado_por, revisado_por)

                        ruta_local = guardar_copia_local_opcional(data_pdf, nom_pdf)
                        guardar_copia_local_opcional(data_excel, nom_xlsx)
                        guardar_copia_local_opcional(data_word, nom_docx)

                        if ruta_local:
                            st.success(f"📂 Copia local guardada automáticamente en: `{ruta_local}`")

                        col_exp1, col_exp2, col_exp3 = st.columns(3)
                        with col_exp1:
                            st.download_button(
                                label="🟢 Descargar Excel (.xlsx)",
                                data=data_excel,
                                file_name=nom_xlsx,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key="dl_xlsx_rep"
                            )
                        with col_exp2:
                            st.download_button(
                                label="🔵 Descargar Word (.docx)",
                                data=data_word,
                                file_name=nom_docx,
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                key="dl_docx_rep"
                            )
                        with col_exp3:
                            st.download_button(
                                label="🔴 Descargar PDF (.pdf)",
                                data=data_pdf,
                                file_name=nom_pdf,
                                mime="application/pdf",
                                key="dl_pdf_rep"
                            )
                    except Exception as e:
                        st.error(f"Error técnico al construir los archivos del reporte: {e}")

# --- MÓDULO 6: RESPALDO Y RESTAURACIÓN (BACKUP) ---
elif menu == "6. Respaldo y Restauración (Backup)":
    st.title("6. Respaldo y Restauración de Base de Datos (Backup)")
    st.caption("Genere copias de seguridad completas del proyecto o restaure un respaldo previo en Google Sheets.")

    st.subheader("1️⃣ Generar y Descargar Respaldo (Backup)")
    st.write(
        "Este respaldo crea un archivo Excel maestro llamado **`bakcup planilla`** (`.xlsx`) con las tres bases de datos completas del sistema: "
        "**Empresas**, **Empleados** y **Datos_Historicos**."
    )

    try:
        df_emp_stat = cargar_datos_gsheet("Empresas", COLUMNAS_EMPRESAS)
        df_empld_stat = cargar_datos_gsheet("Empleados", COLUMNAS_EMPLEADOS)
        df_hist_stat = cargar_datos_gsheet("Datos_Historicos", COLUMNAS_HISTORICO)

        c1, c2, c3 = st.columns(3)
        c1.metric("Empresas a respaldar", len(df_emp_stat))
        c2.metric("Empleados a respaldar", len(df_empld_stat))
        c3.metric("Registros Históricos a respaldar", len(df_hist_stat))

        fecha_hoy_str = datetime.datetime.now().strftime("%Y%m%d_%H%M")
        nombre_bkp = f"bakcup planilla_{fecha_hoy_str}.xlsx"
        bytes_backup = generar_backup_completo_excel()

        ruta_bkp_local = guardar_copia_local_opcional(bytes_backup, nombre_bkp, subcarpeta="bakcup planilla")
        if ruta_bkp_local:
            st.success(f"💾 Copia de seguridad guardada automáticamente en: `{ruta_bkp_local}`")

        st.download_button(
            label="📦 Descargar Archivo 'bakcup planilla' (.xlsx)",
            data=bytes_backup,
            file_name=nombre_bkp,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="btn_descargar_backup_directo"
        )
    except Exception as e:
        st.error(f"Error técnico al preparar el respaldo: {e}")

    st.markdown("---")

    st.subheader("2️⃣ Restaurar Respaldo Automáticamente en el Proyecto")
    st.error(
        "⚠️ **ADVERTENCIA CRÍTICA:** Al subir y restaurar un archivo de respaldo, **ESTO BORRARÁ Y SOBRESCRIBIRÁ TODOS LOS ARCHIVOS Y REGISTROS ELABORADOS ANTES DEL BACKUP** en Google Sheets (Empresas, Empleados y Datos Históricos)."
    )

    archivo_bkp = st.file_uploader(
        "Seleccione el archivo de respaldo (.xlsx) generado previamente por este sistema:",
        type=["xlsx"],
        key="uploader_backup"
    )

    if archivo_bkp is not None:
        confirmar_borrado = st.checkbox(
            "Confirmo que entiendo que esta acción borrará los archivos y registros elaborados antes del backup.",
            key="chk_confirmar_backup"
        )

        if st.button("🔄 Restaurar Respaldo en el Proyecto", disabled=not confirmar_borrado, type="primary"):
            with st.spinner("Restaurando bases de datos en Google Sheets..."):
                exito, mensaje = restaurar_backup_completo_excel(archivo_bkp)
                if exito:
                    st.success(f"✅ {mensaje}")
                    st.balloons()
                else:
                    st.error(f"❌ {mensaje}")
