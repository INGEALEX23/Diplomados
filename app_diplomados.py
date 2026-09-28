import streamlit as st
import pandas as pd
import sqlite3
from datetime import date
import io
import os

# ReportLab para PDF
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

# openpyxl para Excel
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
try:
    from openpyxl.drawing.image import Image as OpenpyxlImage
except ImportError:
    OpenpyxlImage = None

st.set_page_config(
    page_title="Control Escolar - Posgrado e Ingeniería UAdeC",
    layout="wide",
    page_icon="🎓"
)

PATH_LOGO_UADEC = "logo_uadec.png"
PATH_LOGO_POSGRADO = "logo_posgrado.png"
DB_FILE = "diplomados.db"

# ==========================================
# 1. BASE DE DATOS LOCAL MULTI-GRUPO
# ==========================================
def init_db():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    
    # 1. Grupos (Completamente limpio, sin valores por defecto)
    c.execute('''
        CREATE TABLE IF NOT EXISTS grupos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            diplomado TEXT NOT NULL,
            grupo TEXT NOT NULL,
            docente TEXT NOT NULL
        )
    ''')

    # 2. Alumnos
    c.execute('''
        CREATE TABLE IF NOT EXISTS alumnos (
            matricula TEXT,
            grupo_id INTEGER,
            nombre TEXT NOT NULL,
            carrera TEXT,
            telefono TEXT,
            pago_1 INTEGER DEFAULT 0,
            pago_2 INTEGER DEFAULT 0,
            pago_completo INTEGER DEFAULT 0,
            PRIMARY KEY (matricula, grupo_id)
        )
    ''')

    # 3. Bitácora de Horas
    c.execute('''
        CREATE TABLE IF NOT EXISTS bitacora_horas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            grupo_id INTEGER,
            fecha TEXT NOT NULL,
            horas REAL NOT NULL,
            tema TEXT
        )
    ''')

    # 4. Asistencias
    c.execute('''
        CREATE TABLE IF NOT EXISTS asistencias (
            grupo_id INTEGER,
            fecha TEXT NOT NULL,
            matricula TEXT NOT NULL,
            presente INTEGER DEFAULT 1,
            PRIMARY KEY (grupo_id, fecha, matricula)
        )
    ''')

    conn.commit()
    conn.close()

init_db()

# Consultas auxiliares
def get_grupos():
    conn = sqlite3.connect(DB_FILE)
    df = pd.read_sql_query("SELECT id, diplomado, grupo, docente FROM grupos ORDER BY id ASC", conn)
    conn.close()
    return df

def add_grupo(diplomado, grupo, docente):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("INSERT INTO grupos (diplomado, grupo, docente) VALUES (?, ?, ?)",
              (diplomado.strip().upper(), grupo.strip().upper(), docente.strip().upper()))
    conn.commit()
    nuevo_id = c.lastrowid
    conn.close()
    return nuevo_id

def delete_grupo(grupo_id):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("DELETE FROM grupos WHERE id = ?", (grupo_id,))
    c.execute("DELETE FROM alumnos WHERE grupo_id = ?", (grupo_id,))
    c.execute("DELETE FROM bitacora_horas WHERE grupo_id = ?", (grupo_id,))
    c.execute("DELETE FROM asistencias WHERE grupo_id = ?", (grupo_id,))
    conn.commit()
    conn.close()

def get_alumnos_df(grupo_id):
    conn = sqlite3.connect(DB_FILE)
    df = pd.read_sql_query(
        "SELECT matricula, nombre, carrera, telefono, pago_1, pago_2, pago_completo FROM alumnos WHERE grupo_id = ? ORDER BY nombre ASC",
        conn, params=(grupo_id,)
    )
    conn.close()
    df["pago_1"] = df["pago_1"].astype(bool)
    df["pago_2"] = df["pago_2"].astype(bool)
    df["pago_completo"] = df["pago_completo"].astype(bool)
    return df

def save_alumno(matricula, grupo_id, nombre, carrera, telefono, p1, p2, p_comp):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute('''
        INSERT OR REPLACE INTO alumnos (matricula, grupo_id, nombre, carrera, telefono, pago_1, pago_2, pago_completo)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    ''', (
        str(matricula).strip().upper(),
        grupo_id,
        str(nombre).strip().upper(),
        str(carrera).strip().upper(),
        str(telefono).strip().upper(),
        1 if p1 else 0,
        1 if p2 else 0,
        1 if p_comp else 0
    ))
    conn.commit()
    conn.close()

def delete_alumno(matricula, grupo_id):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("DELETE FROM alumnos WHERE matricula = ? AND grupo_id = ?", (matricula, grupo_id))
    c.execute("DELETE FROM asistencias WHERE matricula = ? AND grupo_id = ?", (matricula, grupo_id))
    conn.commit()
    conn.close()

def save_hora(grupo_id, fecha, horas, tema):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("INSERT INTO bitacora_horas (grupo_id, fecha, horas, tema) VALUES (?, ?, ?, ?)",
              (grupo_id, str(fecha), float(horas), str(tema).strip().upper()))
    conn.commit()
    conn.close()

def get_horas_df(grupo_id):
    conn = sqlite3.connect(DB_FILE)
    df = pd.read_sql_query(
        "SELECT id, fecha, horas, tema FROM bitacora_horas WHERE grupo_id = ? ORDER BY fecha DESC, id DESC",
        conn, params=(grupo_id,)
    )
    conn.close()
    return df

def save_asistencia_fecha(grupo_id, fecha, dict_asistencia):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    for mat, pres in dict_asistencia.items():
        c.execute('''
            INSERT OR REPLACE INTO asistencias (grupo_id, fecha, matricula, presente)
            VALUES (?, ?, ?, ?)
        ''', (grupo_id, str(fecha), str(mat), 1 if pres else 0))
    conn.commit()
    conn.close()

def get_asistencia_fecha(grupo_id, fecha):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("SELECT matricula, presente FROM asistencias WHERE grupo_id = ? AND fecha = ?", (grupo_id, str(fecha)))
    res = {row[0]: bool(row[1]) for row in c.fetchall()}
    conn.close()
    return res

# ==========================================
# 2. GENERADORES DE ARCHIVOS (PDF Y EXCEL)
# ==========================================
def generar_reporte_pdf(grupo_id, nombre_dip, grupo, docente):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=letter,
        rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36
    )
    story = []
    styles = getSampleStyleSheet()

    logo_izq = ""
    logo_der = ""
    if os.path.exists(PATH_LOGO_UADEC):
        try:
            logo_izq = RLImage(PATH_LOGO_UADEC, width=70, height=70)
        except Exception:
            logo_izq = ""
    if os.path.exists(PATH_LOGO_POSGRADO):
        try:
            logo_der = RLImage(PATH_LOGO_POSGRADO, width=70, height=70)
        except Exception:
            logo_der = ""

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

    cell_head = ParagraphStyle('CH', parent=styles['Normal'], fontSize=8.5, fontName='Helvetica-Bold', textColor=colors.whitesmoke, alignment=1)
    cell_body = ParagraphStyle('CB', parent=styles['Normal'], fontSize=8, leading=10)

    data = [[
        Paragraph("#", cell_head),
        Paragraph("MATRÍCULA", cell_head),
        Paragraph("NOMBRE COMPLETO", cell_head),
        Paragraph("CARRERA / ESPECIALIDAD", cell_head),
        Paragraph("TELÉFONO", cell_head)
    ]]

    df_al = get_alumnos_df(grupo_id)
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

def generar_excel(grupo_id, con_pagos=False):
    wb = Workbook()
    ws = wb.active
    ws.title = "LISTA OFICIAL"
    ws.views.sheetView[0].showGridLines = True

    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("SELECT diplomado, grupo, docente FROM grupos WHERE id = ?", (grupo_id,))
    res_grupo = c.fetchone()
    conn.close()

    nom_dip = res_grupo[0] if res_grupo else ""
    nom_gpo = res_grupo[1] if res_grupo else ""
    nom_doc = res_grupo[2] if res_grupo else ""

    df_al = get_alumnos_df(grupo_id)
    col_fin_letra = "H" if con_pagos else "E"
    col_fin_num = 8 if con_pagos else 5

    # Insertar logos si existen
    if OpenpyxlImage is not None:
        if os.path.exists(PATH_LOGO_UADEC):
            try:
                img_uadec = OpenpyxlImage(PATH_LOGO_UADEC)
                img_uadec.width = 75
                img_uadec.height = 75
                ws.add_image(img_uadec, "A1")
            except Exception:
                pass

        if os.path.exists(PATH_LOGO_POSGRADO):
            try:
                img_pos = OpenpyxlImage(PATH_LOGO_POSGRADO)
                img_pos.width = 75
                img_pos.height = 75
                ws.add_image(img_pos, f"{col_fin_letra}1")
            except Exception:
                pass

    # Membrete
    col_centro = chr(64 + col_fin_num - 1)
    ws.merge_cells(f"B1:{col_centro}1")
    ws["B1"] = "UNIVERSIDAD AUTÓNOMA DE COAHUILA"
    ws["B1"].font = Font(name="Calibri", size=13, bold=True, color="1A365D")
    ws["B1"].alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells(f"B2:{col_centro}2")
    ws["B2"] = "FACULTAD DE INGENIERÍA - COORDINACIÓN DE POSGRADO"
    ws["B2"].font = Font(name="Calibri", size=10, bold=True, color="4A5568")
    ws["B2"].alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells(f"B3:{col_centro}3")
    ws["B3"] = "REPORTE OFICIAL DE PARTICIPANTES" if not con_pagos else "CONTROL INTERNO Y REPORTE DE PAGOS"
    ws["B3"].font = Font(name="Calibri", size=9, bold=False, color="718096")
    ws["B3"].alignment = Alignment(horizontal="center", vertical="center")

    ws.row_dimensions[1].height = 24
    ws.row_dimensions[2].height = 18
    ws.row_dimensions[3].height = 16
    ws.row_dimensions[4].height = 10

    # Tarjeta de Datos
    fill_meta = PatternFill(start_color="EDF2F7", end_color="EDF2F7", fill_type="solid")
    font_meta = Font(name="Calibri", size=9, color="1A202C")
    thin_border_side = Side(style='thin', color="CBD5E0")
    border_meta = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)

    ws["A5"] = f"DIPLOMADO: {nom_dip}"
    ws["A5"].font = font_meta
    ws["D5"] = f"GRUPO / GENERACIÓN: {nom_gpo}"
    ws["D5"].font = font_meta

    ws["A6"] = f"INSTRUCTOR: {nom_doc}"
    ws["A6"].font = font_meta
    ws["D6"] = f"FECHA DE EMISIÓN: {date.today().strftime('%d/%m/%Y')}"
    ws["D6"].font = font_meta

    ws.merge_cells("A5:C5")
    ws.merge_cells(f"D5:{col_fin_letra}5")
    ws.merge_cells("A6:C6")
    ws.merge_cells(f"D6:{col_fin_letra}6")

    for r in range(5, 7):
        ws.row_dimensions[r].height = 20
        for col_idx in range(1, col_fin_num + 1):
            cell = ws.cell(row=r, column=col_idx)
            cell.fill = fill_meta
            cell.border = border_meta
            cell.alignment = Alignment(vertical="center", indent=1)

    ws.row_dimensions[7].height = 12

    headers = ["#", "MATRÍCULA", "NOMBRE COMPLETO", "CARRERA / ESPECIALIDAD", "TELÉFONO"]
    if con_pagos:
        headers.extend(["1ER PAGO", "2DO PAGO", "PAGO COMPLETO"])

    header_fill = PatternFill(start_color="1A365D", end_color="1A365D", fill_type="solid")
    header_font = Font(name="Calibri", size=9.5, bold=True, color="FFFFFF")
    header_border = Border(
        left=Side(style='thin', color="4A5568"),
        right=Side(style='thin', color="4A5568"),
        top=Side(style='thin', color="1A365D"),
        bottom=Side(style='medium', color="1A365D")
    )

    ws.row_dimensions[8].height = 24
    for c_idx, h_text in enumerate(headers, 1):
        cell = ws.cell(row=8, column=c_idx, value=h_text)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = header_border

    body_font = Font(name="Calibri", size=9, color="2D3748")
    zebra_fill = PatternFill(start_color="F7FAFC", end_color="F7FAFC", fill_type="solid")
    white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    cell_border = Border(
        left=Side(style='thin', color="E2E8F0"),
        right=Side(style='thin', color="E2E8F0"),
        top=Side(style='thin', color="E2E8F0"),
        bottom=Side(style='thin', color="E2E8F0")
    )

    fila_actual = 9
    for idx, row in df_al.iterrows():
        ws.row_dimensions[fila_actual].height = 20
        fill_row = zebra_fill if idx % 2 == 1 else white_fill
        valores = [idx + 1, row["matricula"], row["nombre"], row["carrera"], row["telefono"]]
        if con_pagos:
            valores.extend([
                "PAGADO" if row["pago_1"] else "PENDIENTE",
                "PAGADO" if row["pago_2"] else "PENDIENTE",
                "LIQUIDADO" if row["pago_completo"] else "PENDIENTE"
            ])

        for c_idx, val in enumerate(valores, 1):
            cell = ws.cell(row=fila_actual, column=c_idx, value=val)
            cell.font = body_font
            cell.fill = fill_row
            cell.border = cell_border
            if c_idx in [1, 2] or (con_pagos and c_idx >= 6):
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        fila_actual += 1

    anchos = {"A": 6, "B": 15, "C": 38, "D": 32, "E": 16}
    if con_pagos:
        anchos["F"] = 14
        anchos["G"] = 14
        anchos["H"] = 16

    for col_letter, width in anchos.items():
        ws.column_dimensions[col_letter].width = width

    if con_pagos:
        df_h = get_horas_df(grupo_id)
        if not df_h.empty:
            ws_h = wb.create_sheet(title="BITÁCORA DE HORAS")
            ws_h.views.sheetView[0].showGridLines = True
            headers_h = ["FECHA", "HORAS IMPARTIDAS", "TEMA / ACTIVIDAD ABORDADA"]
            ws_h.row_dimensions[1].height = 24
            for c_idx, h_text in enumerate(headers_h, 1):
                cell = ws_h.cell(row=1, column=c_idx, value=h_text)
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = header_border

            for h_idx, h_row in df_h.iterrows():
                f_idx = h_idx + 2
                ws_h.row_dimensions[f_idx].height = 19
                fill_row = zebra_fill if h_idx % 2 == 1 else white_fill
                c_f = ws_h.cell(row=f_idx, column=1, value=str(h_row["fecha"]))
                c_h = ws_h.cell(row=f_idx, column=2, value=float(h_row["horas"]))
                c_t = ws_h.cell(row=f_idx, column=3, value=str(h_row["tema"]))
                for c in [c_f, c_h, c_t]:
                    c.font = body_font
                    c.fill = fill_row
                    c.border = cell_border
                c_f.alignment = Alignment(horizontal="center", vertical="center")
                c_h.alignment = Alignment(horizontal="center", vertical="center")
                c_t.alignment = Alignment(horizontal="left", vertical="center", indent=1)

            ws_h.column_dimensions["A"].width = 15
            ws_h.column_dimensions["B"].width = 18
            ws_h.column_dimensions["C"].width = 50

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# ==========================================
# 3. INTERFAZ DE USUARIO
# ==========================================
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

# Consultar grupos registrados
df_grupos = get_grupos()

# PANTALLA DE INICIO SI NO HAY GRUPOS CREADOS
if df_grupos.empty:
    st.info("👋 BIENVENIDO. NO HAY NINGÚN DIPLOMADO O GRUPO REGISTRADO.")
    st.subheader("➕ CREAR EL PRIMER GRUPO DE DIPLOMADO")
    with st.form("form_primer_grupo"):
        nvo_dip = st.text_input("NOMBRE DEL DIPLOMADO").upper()
        nvo_gpo = st.text_input("GRUPO / CLAVE (EJ. GPO. 2026-I)").upper()
        nvo_doc = st.text_input("DOCENTE RESPONSABLE").upper()
        if st.form_submit_button("REGISTRAR Y COMENZAR"):
            if nvo_dip and nvo_gpo:
                add_grupo(nvo_dip, nvo_gpo, nvo_doc)
                st.success("DIPLOMADO REGISTRADO CORRECTAMENTE.")
                st.rerun()
            else:
                st.error("EL NOMBRE DEL DIPLOMADO Y EL GRUPO SON OBLIGATORIOS.")
    st.stop()

# Si ya hay grupos registrados, se activa la barra lateral
with st.sidebar:
    st.header("📚 GRUPO ACTIVO")
    opciones_grupos = {
        row["id"]: f"{row['grupo']} - {row['diplomado']}"
        for _, row in df_grupos.iterrows()
    }

    grupo_seleccionado_id = st.selectbox(
        "SELECCIONA EL DIPLOMADO / GRUPO",
        options=list(opciones_grupos.keys()),
        format_func=lambda x: opciones_grupos[x]
    )

    info_grupo = df_grupos[df_grupos["id"] == grupo_seleccionado_id].iloc[0]
    dip_nombre = info_grupo["diplomado"]
    dip_grupo = info_grupo["grupo"]
    dip_docente = info_grupo["docente"]

    st.info(f"**Diplomado:** {dip_nombre}\n\n**Grupo:** {dip_grupo}\n\n**Docente:** {dip_docente}")

    df_horas_tot = get_horas_df(grupo_seleccionado_id)
    tot_h = df_horas_tot["horas"].sum() if not df_horas_tot.empty else 0.0
    st.metric("HORAS ACUMULADAS", f"{tot_h:.1f} HRS")

    st.divider()
    st.subheader("📥 DESCARGAS Y REPORTES")

    pdf_bytes = generar_reporte_pdf(grupo_seleccionado_id, dip_nombre, dip_grupo, dip_docente)
    st.download_button(
        label="📄 DESCARGAR PDF OFICIAL (SIN PAGOS)",
        data=pdf_bytes,
        file_name=f"LISTA_OFICIAL_{dip_grupo.replace(' ', '_')}.pdf",
        mime="application/pdf",
        use_container_width=True
    )

    excel_sin = generar_excel(grupo_seleccionado_id, con_pagos=False)
    st.download_button(
        label="📊 DESCARGAR EXCEL (SIN PAGOS)",
        data=excel_sin,
        file_name=f"ALUMNOS_{dip_grupo.replace(' ', '_')}_OFICIAL.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

    excel_con = generar_excel(grupo_seleccionado_id, con_pagos=True)
    st.download_button(
        label="💰 DESCARGAR EXCEL (CON PAGOS Y HORAS)",
        data=excel_con,
        file_name=f"CONTROL_{dip_grupo.replace(' ', '_')}_PAGOS.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

    st.divider()
    with st.expander("➕ CREAR OTRO GRUPO / DIPLOMADO"):
        with st.form("form_nuevo_grupo"):
            nvo_dip = st.text_input("NOMBRE DEL DIPLOMADO").upper()
            nvo_gpo = st.text_input("GRUPO / CLAVE").upper()
            nvo_doc = st.text_input("DOCENTE RESPONSABLE", value=dip_docente).upper()
            if st.form_submit_button("CREAR GRUPO"):
                if nvo_dip and nvo_gpo:
                    add_grupo(nvo_dip, nvo_gpo, nvo_doc)
                    st.success("GRUPO CREADO EXITOSAMENTE.")
                    st.rerun()
                else:
                    st.error("NOMBRE Y GRUPO SON REQUERIDOS.")

    with st.expander("🗑️ ELIMINAR ESTE GRUPO"):
        st.caption(f"¿Deseas eliminar '{dip_grupo}' y todos sus datos?")
        if st.button("CONFIRMAR Y BORRAR ESTE GRUPO"):
            delete_grupo(grupo_seleccionado_id)
            st.warning("GRUPO ELIMINADO.")
            st.rerun()

# Pestañas de trabajo para el grupo seleccionado
tab_alumnos, tab_asistencia, tab_horas = st.tabs([
    f"📋 ALUMNOS Y PAGOS ({dip_grupo})",
    f"✅ PASE DE LISTA",
    f"⏱️ BITÁCORA DE HORAS"
])

# --- TAB 1: ALUMNOS Y PAGOS ---
with tab_alumnos:
    df_al = get_alumnos_df(grupo_seleccionado_id)

    st.subheader(f"LISTADO DE ESTUDIANTES - {dip_grupo}")
    if not df_al.empty:
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
            key=f"editor_{grupo_seleccionado_id}"
        )

        c_save, c_del = st.columns([2, 2])
        if c_save.button("💾 GUARDAR MODIFICACIONES DE TABLA", use_container_width=True):
            for _, r in edited_data.iterrows():
                save_alumno(
                    r["MATRÍCULA"],
                    grupo_seleccionado_id,
                    r["NOMBRE"],
                    r["CARRERA"],
                    r["TELÉFONO"],
                    r["1ER PAGO"],
                    r["2DO PAGO"],
                    r["PAGO COMPLETO"]
                )
            st.success("DATOS GUARDADOS EN DISCO.")
            st.rerun()

        with c_del:
            with st.popover("🗑️ ELIMINAR UN ALUMNO"):
                mat_del = st.selectbox("MATRÍCULA A ELIMINAR", df_al["matricula"].tolist())
                if st.button("CONFIRMAR ELIMINACIÓN"):
                    delete_alumno(mat_del, grupo_seleccionado_id)
                    st.warning(f"ALUMNO {mat_del} ELIMINADO.")
                    st.rerun()
    else:
        st.info("NO HAY ALUMNOS EN ESTE GRUPO AÚN.")

    st.divider()

    with st.expander("➕ REGISTRAR NUEVO ALUMNO A ESTE GRUPO", expanded=df_al.empty):
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

            if st.form_submit_button("GUARDAR ALUMNO"):
                if n_mat and n_nom:
                    save_alumno(n_mat, grupo_seleccionado_id, n_nom, n_car, n_tel, p1, p2, p_comp)
                    st.success(f"ALUMNO {n_nom} REGISTRADO EN {dip_grupo}.")
                    st.rerun()
                else:
                    st.error("LA MATRÍCULA Y EL NOMBRE SON OBLIGATORIOS.")

# --- TAB 2: PASE DE LISTA ---
with tab_asistencia:
    st.subheader(f"PASE DE LISTA - {dip_grupo}")
    col_fa, _ = st.columns([1, 2])
    f_asis = col_fa.date_input("FECHA DE LA SESIÓN", value=date.today())
    str_f = str(f_asis)

    asis_guardadas = get_asistencia_fecha(grupo_seleccionado_id, str_f)
    df_al_asis = get_alumnos_df(grupo_seleccionado_id)

    if not df_al_asis.empty:
        with st.form("form_pase_lista"):
            st.write(f"REGISTRO PARA LA SESIÓN: **{str_f}**")
            nuevos_estados = {}
            for _, al in df_al_asis.iterrows():
                val_previo = asis_guardadas.get(al["matricula"], True)
                col_n, col_c = st.columns([3, 1])
                col_n.write(f"**{al['nombre']}** ({al['matricula']})")
                nuevos_estados[al["matricula"]] = col_c.checkbox("PRESENTE", value=val_previo, key=f"att_{grupo_seleccionado_id}_{str_f}_{al['matricula']}")

            if st.form_submit_button("💾 GUARDAR PASE DE LISTA"):
                save_asistencia_fecha(grupo_seleccionado_id, str_f, nuevos_estados)
                st.success(f"ASISTENCIAS DEL {str_f} GUARDADAS.")
    else:
        st.info("ESTE GRUPO NO TIENE ALUMNOS REGISTRADOS TODAVÍA.")

# --- TAB 3: CONTROL DE HORAS ---
with tab_horas:
    st.subheader(f"BITÁCORA DOCENTE DE HORAS - {dip_grupo}")
    col_h1, col_h2 = st.columns([1, 1])

    with col_h1:
        with st.form("form_nueva_hora"):
            st.write("**REGISTRAR SESIÓN IMPARTIDA**")
            f_h = st.date_input("FECHA", value=date.today(), key=f"f_hora_{grupo_seleccionado_id}")
            n_hrs = st.number_input("HORAS DADAS", min_value=0.5, max_value=12.0, value=2.0, step=0.5)
            n_tema = st.text_input("TEMA / CONTENIDO ABORDADO").upper()

            if st.form_submit_button("REGISTRAR HORAS"):
                if n_tema:
                    save_hora(grupo_seleccionado_id, f_h, n_hrs, n_tema)
                    st.success("HORAS REGISTRADAS EXITOSAMENTE.")
                    st.rerun()
                else:
                    st.error("INGRESE EL TEMA DE LA CLASE.")

    with col_h2:
        st.write("**HISTORIAL DEL GRUPO**")
        df_hist = get_horas_df(grupo_seleccionado_id)
        if not df_hist.empty:
            df_hist_view = df_hist.rename(columns={"fecha": "FECHA", "horas": "HORAS", "tema": "TEMA"})
            st.dataframe(df_hist_view[["FECHA", "HORAS", "TEMA"]], use_container_width=True)
        else:
            st.caption("AÚN NO HAY HORAS REGISTRADAS EN ESTE GRUPO.")
