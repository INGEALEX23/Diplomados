import streamlit as st
import pandas as pd
import sqlite3
from datetime import date
import io
import os
from PIL import Image

# ReportLab para el PDF
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

# Configuración general
st.set_page_config(
    page_title="Control Escolar - Posgrado e Ingeniería UAdeC",
    layout="wide",
    page_icon="🎓"
)

# Nombres de archivos de logotipos
PATH_LOGO_UADEC = "logo_uadec.png"
PATH_LOGO_POSGRADO = "logo_posgrado.png"

# ==========================================
# 1. BASE DE DATOS LOCAL (PERSISTENCIA SQLITE)
# ==========================================
DB_FILE = "diplomados.db"

def init_db():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    # Tabla de Alumnos
    c.execute('''
        CREATE TABLE IF NOT EXISTS alumnos (
            matricula TEXT PRIMARY KEY,
            nombre TEXT NOT NULL,
            carrera TEXT,
            telefono TEXT,
            pago_1 INTEGER DEFAULT 0,
            pago_2 INTEGER DEFAULT 0,
            pago_completo INTEGER DEFAULT 0
        )
    ''')
    # Tabla de Bitácora de Horas
    c.execute('''
        CREATE TABLE IF NOT EXISTS bitacora_horas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            fecha TEXT NOT NULL,
            horas REAL NOT NULL,
            tema TEXT
        )
    ''')
    # Tabla de Asistencias
    c.execute('''
        CREATE TABLE IF NOT EXISTS asistencias (
            fecha TEXT NOT NULL,
            matricula TEXT NOT NULL,
            presente INTEGER DEFAULT 1,
            PRIMARY KEY (fecha, matricula)
        )
    ''')
    conn.commit()
    conn.close()

init_db()

def get_alumnos_df():
    conn = sqlite3.connect(DB_FILE)
    df = pd.read_sql_query("SELECT matricula, nombre, carrera, telefono, pago_1, pago_2, pago_completo FROM alumnos ORDER BY nombre ASC", conn)
    conn.close()
    df["pago_1"] = df["pago_1"].astype(bool)
    df["pago_2"] = df["pago_2"].astype(bool)
    df["pago_completo"] = df["pago_completo"].astype(bool)
    return df

def save_alumno(matricula, nombre, carrera, telefono, p1, p2, p_comp):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute('''
        INSERT OR REPLACE INTO alumnos (matricula, nombre, carrera, telefono, pago_1, pago_2, pago_completo)
        VALUES (?, ?, ?, ?, ?, ?, ?)
    ''', (
        str(matricula).strip().upper(),
        str(nombre).strip().upper(),
        str(carrera).strip().upper(),
        str(telefono).strip().upper(),
        1 if p1 else 0,
        1 if p2 else 0,
        1 if p_comp else 0
    ))
    conn.commit()
    conn.close()

def delete_alumno(matricula):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("DELETE FROM alumnos WHERE matricula = ?", (matricula,))
    c.execute("DELETE FROM asistencias WHERE matricula = ?", (matricula,))
    conn.commit()
    conn.close()

def save_hora(fecha, horas, tema):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("INSERT INTO bitacora_horas (fecha, horas, tema) VALUES (?, ?, ?)",
              (str(fecha), float(horas), str(tema).strip().upper()))
    conn.commit()
    conn.close()

def get_horas_df():
    conn = sqlite3.connect(DB_FILE)
    df = pd.read_sql_query("SELECT id, fecha, horas, tema FROM bitacora_horas ORDER BY fecha DESC, id DESC", conn)
    conn.close()
    return df

def save_asistencia_fecha(fecha, dict_asistencia):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    for mat, pres in dict_asistencia.items():
        c.execute('''
            INSERT OR REPLACE INTO asistencias (fecha, matricula, presente)
            VALUES (?, ?, ?)
        ''', (str(fecha), str(mat), 1 if pres else 0))
    conn.commit()
    conn.close()

def get_asistencia_fecha(fecha):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("SELECT matricula, presente FROM asistencias WHERE fecha = ?", (str(fecha),))
    res = {row[0]: bool(row[1]) for row in c.fetchall()}
    conn.close()
    return res

# ==========================================
# 2. GENERADORES DE ARCHIVOS (PDF Y EXCEL)
# ==========================================
def generar_reporte_pdf(nombre_dip, grupo, docente):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=letter,
        rightMargin=36,
        leftMargin=36,
        topMargin=36,
        bottomMargin=36
    )
    story = []
    styles = getSampleStyleSheet()

    # Encabezado institucional con logos duales
    logo_izq = ""
    logo_der = ""
    if os.path.exists(PATH_LOGO_UADEC):
        logo_izq = RLImage(PATH_LOGO_UADEC, width=70, height=70)
    if os.path.exists(PATH_LOGO_POSGRADO):
        logo_der = RLImage(PATH_LOGO_POSGRADO, width=70, height=70)

    header_text = Paragraph(
        f"<font size=11><b>UNIVERSIDAD AUTÓNOMA DE COAHUILA</b></font><br/>"
        f"<font size=9 color='#2D3748'>FACULTAD DE INGENIERÍA - COORDINACIÓN DE POSGRADO</font><br/>"
        f"<font size=8 color='#718096'>REPORTE OFICIAL DE PARTICIPANTES</font>",
        ParagraphStyle('HeaderTxt', alignment=1, leading=13)
    )

    t_header = Table([[logo_izq, header_text, logo_der]], colWidths=[80, 380, 80])
    t_header.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('ALIGN', (0,0), (0,0), 'LEFT'),
        ('ALIGN', (2,0), (2,0), 'RIGHT'),
        ('BOTTOMPADDING', (0,0), (-1,-1), 10),
    ]))
    story.append(t_header)
    story.append(Spacer(1, 10))

    # Datos generales
    meta_style = ParagraphStyle('Meta', parent=styles['Normal'], fontSize=8.5, leading=12)
    meta_data = [
        [
            Paragraph(f"<b>DIPLOMADO:</b> {nombre_dip.upper()}", meta_style),
            Paragraph(f"<b>GRUPO / GENERACIÓN:</b> {grupo.upper()}", meta_style)
        ],
        [
            Paragraph(f"<b>INSTRUCTOR:</b> {docente.upper()}", meta_style),
            Paragraph(f"<b>FECHA DE EMISIÓN:</b> {date.today().strftime('%d/%m/%Y')}", meta_style)
        ]
    ]
    t_meta = Table(meta_data, colWidths=[270, 270])
    t_meta.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#EDF2F7")),
        ('BOX', (0,0), (-1,-1), 0.5, colors.HexColor("#CBD5E0")),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 5),
    ]))
    story.append(t_meta)
    story.append(Spacer(1, 15))

    # Tabla de Alumnos (SOLO DATOS ESCOLARES, NADA DE PAGOS NI HORAS)
    cell_head = ParagraphStyle('CH', parent=styles['Normal'], fontSize=8.5, fontName='Helvetica-Bold', textColor=colors.whitesmoke, alignment=1)
    cell_body = ParagraphStyle('CB', parent=styles['Normal'], fontSize=8, leading=10)

    data = [[
        Paragraph("#", cell_head),
        Paragraph("MATRÍCULA", cell_head),
        Paragraph("NOMBRE COMPLETO", cell_head),
        Paragraph("CARRERA / ESPECIALIDAD", cell_head),
        Paragraph("TELÉFONO", cell_head)
    ]]

    df_al = get_alumnos_df()
    for idx, row in df_al.iterrows():
        data.append([
            Paragraph(str(idx + 1), cell_body),
            Paragraph(str(row["matricula"]), cell_body),
            Paragraph(str(row["nombre"]), cell_body),
            Paragraph(str(row["carrera"]), cell_body),
            Paragraph(str(row["telefono"]), cell_body)
        ])

    t_alumnos = Table(data, colWidths=[25, 75, 210, 150, 80], repeatRows=1)
    t_alumnos.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#1A365D")),
        ('ALIGN', (0, 0), (0, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E0")),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor("#F7FAFC")]),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
    ]))
    story.append(t_alumnos)

    doc.build(story)
    buffer.seek(0)
    return buffer

def generar_excel(con_pagos=False):
    output = io.BytesIO()
    df = get_alumnos_df()
    
    if not con_pagos:
        # Solo escolar
        df_export = df[["matricula", "nombre", "carrera", "telefono"]].copy()
        df_export.columns = ["MATRÍCULA", "NOMBRE COMPLETO", "CARRERA", "TELÉFONO"]
    else:
        # Completo con pagos
        df_export = df.copy()
        df_export["pago_1"] = df_export["pago_1"].apply(lambda x: "PAGADO" if x else "PENDIENTE")
        df_export["pago_2"] = df_export["pago_2"].apply(lambda x: "PAGADO" if x else "PENDIENTE")
        df_export["pago_completo"] = df_export["pago_completo"].apply(lambda x: "LIQUIDADO" if x else "PENDIENTE")
        df_export.columns = ["MATRÍCULA", "NOMBRE COMPLETO", "CARRERA", "TELÉFONO", "1ER PAGO", "2DO PAGO", "PAGO COMPLETO"]

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_export.to_excel(writer, index=False, sheet_name="ALUMNOS")
        
        # Si es con pagos, agregamos también una pestaña de horas impartidas
        if con_pagos:
            df_h = get_horas_df()
            if not df_h.empty:
                df_h_exp = df_h[["fecha", "horas", "tema"]].copy()
                df_h_exp.columns = ["FECHA", "HORAS IMPARTIDAS", "TEMA / ACTIVIDAD"]
                df_h_exp.to_excel(writer, index=False, sheet_name="BITÁCORA HORAS")

    output.seek(0)
    return output

# ==========================================
# 3. INTERFAZ DE USUARIO (STREAMLIT)
# ==========================================
# Encabezado visual con logos
col_l1, col_t, col_l2 = st.columns([1, 4, 1])

with col_l1:
    if os.path.exists(PATH_LOGO_UADEC):
        st.image(PATH_LOGO_UADEC, use_container_width=True)
    else:
        st.caption("🏛️ UAdeC")

with col_t:
    st.markdown("<h2 style='text-align: center; margin-bottom: 0;'>COORDINACIÓN DE POSGRADO DE INGENIERÍA</h2>", unsafe_allow_html=True)
    st.markdown("<h4 style='text-align: center; color: gray; margin-top: 0;'>SISTEMA DE GESTIÓN Y CONTROL DE DIPLOMADOS</h4>", unsafe_allow_html=True)

with col_l2:
    if os.path.exists(PATH_LOGO_POSGRADO):
        st.image(PATH_LOGO_POSGRADO, use_container_width=True)
    else:
        st.caption("🔬 POSGRADO")

st.divider()

# Barra lateral con Metadatos y Descargas
with st.sidebar:
    st.header("⚙️ CONFIGURACIÓN")
    dip_nombre = st.text_input("NOMBRE DEL DIPLOMADO", value="AUTOMATIZACIÓN INDUSTRIAL Y SISTEMAS FOTOVOLTAICOS").upper()
    dip_grupo = st.text_input("GRUPO / GENERACIÓN", value="GPO. 2026-I").upper()
    dip_docente = st.text_input("DOCENTE RESPONSABLE", value="MTRO. EN ING. OMAR ALEJANDRO JIMÉNEZ NAVARRO").upper()

    df_horas_tot = get_horas_df()
    tot_h = df_horas_tot["horas"].sum() if not df_horas_tot.empty else 0.0
    st.metric("HORAS ACUMULADAS", f"{tot_h:.1f} HRS")

    st.divider()
    st.subheader("📥 DESCARGAS Y REPORTES")

    # 1. Botón PDF
    pdf_bytes = generar_reporte_pdf(dip_nombre, dip_grupo, dip_docente)
    st.download_button(
        label="📄 DESCARGAR PDF OFICIAL (SIN PAGOS)",
        data=pdf_bytes,
        file_name=f"LISTA_OFICIAL_{dip_grupo.replace(' ', '_')}.pdf",
        mime="application/pdf",
        use_container_width=True
    )

    # 2. Botón Excel SIN PAGOS
    excel_sin = generar_excel(con_pagos=False)
    st.download_button(
        label="📊 DESCARGAR EXCEL (SIN PAGOS)",
        data=excel_sin,
        file_name=f"ALUMNOS_{dip_grupo.replace(' ', '_')}_OFICIAL.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

    # 3. Botón Excel CON PAGOS
    excel_con = generar_excel(con_pagos=True)
    st.download_button(
        label="💰 DESCARGAR EXCEL (CON PAGOS Y HORAS)",
        data=excel_con,
        file_name=f"CONTROL_INTERNO_{dip_grupo.replace(' ', '_')}_PAGOS.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

# Pestañas principales
tab_alumnos, tab_asistencia, tab_horas = st.tabs([
    "📋 ALUMNOS Y PAGOS",
    "✅ PASE DE LISTA",
    "⏱️ BITÁCORA DE HORAS"
])

# --- TAB 1: ALUMNOS Y CONTROL FINANCIERO ---
with tab_alumnos:
    df_al = get_alumnos_df()

    st.subheader("LISTADO DE ESTUDIANTES")
    if not df_al.empty:
        # Formateo para editor
        df_display = df_al.rename(columns={
            "matricula": "MATRÍCULA",
            "nombre": "NOMBRE",
            "carrera": "CARRERA",
            "telefono": "TELÉFONO",
            "pago_1": "1ER PAGO",
            "pago_2": "2DO PAGO",
            "pago_completo": "PAGO COMPLETO"
        })

        col_cfg = {
            "MATRÍCULA": st.column_config.TextColumn(required=True),
            "NOMBRE": st.column_config.TextColumn(required=True),
            "CARRERA": st.column_config.TextColumn(),
            "TELÉFONO": st.column_config.TextColumn(),
            "1ER PAGO": st.column_config.CheckboxColumn(),
            "2DO PAGO": st.column_config.CheckboxColumn(),
            "PAGO COMPLETO": st.column_config.CheckboxColumn()
        }

        edited_data = st.data_editor(
            df_display,
            use_container_width=True,
            column_config=col_cfg,
            num_rows="fixed",
            key="editor_grid"
        )

        c_save, c_del = st.columns([2, 2])
        if c_save.button("💾 GUARDAR MODIFICACIONES DE TABLA", use_container_width=True):
            for _, r in edited_data.iterrows():
                save_alumno(
                    r["MATRÍCULA"],
                    r["NOMBRE"],
                    r["CARRERA"],
                    r["TELÉFONO"],
                    r["1ER PAGO"],
                    r["2DO PAGO"],
                    r["PAGO COMPLETO"]
                )
            st.success("DATOS ACTUALIZADOS Y ALMACENADOS EN DISCO.")
            st.rerun()

        with c_del:
            with st.popover("🗑️ ELIMINAR UN ALUMNO"):
                mat_del = st.selectbox("SELECCIONA MATRÍCULA A ELIMINAR", df_al["matricula"].tolist())
                if st.button("CONFIRMAR ELIMINACIÓN"):
                    delete_alumno(mat_del)
                    st.warning(f"ALUMNO {mat_del} ELIMINADO.")
                    st.rerun()
    else:
        st.info("NO HAY ALUMNOS REGISTRADOS EN LA BASE DE DATOS.")

    st.divider()

    with st.expander("➕ REGISTRAR NUEVO ALUMNO", expanded=df_al.empty):
        with st.form("form_nuevo_alumno"):
            c1, c2 = st.columns(2)
            n_mat = c1.text_input("MATRÍCULA").upper()
            n_nom = c2.text_input("NOMBRE COMPLETO (APELLIDOS Y NOMBRES)").upper()
            c3, c4 = st.columns(2)
            n_car = c3.text_input("CARRERA / FACULTAD").upper()
            n_tel = c4.text_input("NÚMERO DE TELÉFONO").upper()

            st.write("**ESTATUS DE PAGO:**")
            cp1, cp2, cp3 = st.columns(3)
            p1 = cp1.checkbox("1ER PAGO")
            p2 = cp2.checkbox("2DO PAGO")
            p_comp = cp3.checkbox("PAGO COMPLETO")

            if st.form_submit_button("GUARDAR EN BASE DE DATOS"):
                if n_mat and n_nom:
                    save_alumno(n_mat, n_nom, n_car, n_tel, p1, p2, p_comp)
                    st.success(f"ALUMNO {n_nom} REGISTRADO CORRECTAMENTE.")
                    st.rerun()
                else:
                    st.error("LA MATRÍCULA Y EL NOMBRE SON OBLIGATORIOS.")

# --- TAB 2: PASE DE LISTA ---
with tab_asistencia:
    st.subheader("PASE DE LISTA DIARIO")
    col_fa, _ = st.columns([1, 2])
    f_asis = col_fa.date_input("FECHA DE LA SESIÓN", value=date.today())
    str_f = str(f_asis)

    asis_guardadas = get_asistencia_fecha(str_f)
    df_al_asis = get_alumnos_df()

    if not df_al_asis.empty:
        with st.form("form_pase_lista"):
            st.write(f"REGISTRO PARA LA SESIÓN: **{str_f}**")
            nuevos_estados = {}
            for _, al in df_al_asis.iterrows():
                val_previo = asis_guardadas.get(al["matricula"], True)
                col_n, col_c = st.columns([3, 1])
                col_n.write(f"**{al['nombre']}** ({al['matricula']})")
                nuevos_estados[al["matricula"]] = col_c.checkbox("PRESENTE", value=val_previo, key=f"att_{str_f}_{al['matricula']}")

            if st.form_submit_button("💾 GUARDAR PASE DE LISTA"):
                save_asistencia_fecha(str_f, nuevos_estados)
                st.success(f"ASISTENCIAS DEL DÍA {str_f} GUARDADAS EN BASE DE DATOS.")
    else:
        st.info("PRIMERO REGISTRE ALUMNOS EN LA PESTAÑA ANTERIOR.")

# --- TAB 3: CONTROL DE HORAS ---
with tab_horas:
    st.subheader("BITÁCORA DOCENTE DE HORAS")
    col_h1, col_h2 = st.columns([1, 1])

    with col_h1:
        with st.form("form_nueva_hora"):
            st.write("**REGISTRAR SESIÓN IMPARTIDA**")
            f_h = st.date_input("FECHA", value=date.today(), key="f_hora")
            n_hrs = st.number_input("HORAS DADAS", min_value=0.5, max_value=12.0, value=2.0, step=0.5)
            n_tema = st.text_input("TEMA / CONTENIDO ABORDADO").upper()

            if st.form_submit_button("REGISTRAR HORAS"):
                if n_tema:
                    save_hora(f_h, n_hrs, n_tema)
                    st.success("HORAS REGISTRADAS EXITOSAMENTE.")
                    st.rerun()
                else:
                    st.error("INGRESE EL TEMA DE LA CLASE.")

    with col_h2:
        st.write("**HISTORIAL REGISTRADO**")
        df_hist = get_horas_df()
        if not df_hist.empty:
            df_hist_view = df_hist.rename(columns={"fecha": "FECHA", "horas": "HORAS", "tema": "TEMA"})
            st.dataframe(df_hist_view[["FECHA", "HORAS", "TEMA"]], use_container_width=True)
        else:
            st.caption("AÚN NO HAY HORAS REGISTRADAS.")
