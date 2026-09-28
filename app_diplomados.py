import streamlit as st
import pandas as pd
from datetime import date
import io
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

# Configuración de página
st.set_page_config(page_title="Gestión de Diplomados", layout="wide", page_icon="🎓")

# Inicialización de estado en memoria (Session State)
if "alumnos" not in st.session_state:
    st.session_state.alumnos = [
        {
            "matricula": "18045001",
            "nombre": "García López Carlos",
            "carrera": "Ing. Mecánica Eléctrica",
            "telefono": "844-123-4567",
            "pago_1": True,
            "pago_2": False,
            "pago_completo": False
        },
        {
            "matricula": "19045032",
            "nombre": "Martínez Ramos Sofía",
            "carrera": "Ing. en Robótica y Sistemas",
            "telefono": "844-987-6543",
            "pago_1": True,
            "pago_2": True,
            "pago_completo": True
        }
    ]

if "bitacora_horas" not in st.session_state:
    st.session_state.bitacora_horas = []

if "asistencias" not in st.session_state:
    st.session_state.asistencias = {}

# --- FUNCIÓN GENERADORA DEL REPORTE PDF ---
def generar_reporte_pdf(nombre_diplomado, grupo_nombre, instructores):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=letter,
        rightMargin=36,
        leftMargin=36,
        topMargin=40,
        bottomMargin=36
    )
    story = []
    styles = getSampleStyleSheet()

    # Estilos tipográficos
    title_style = ParagraphStyle(
        'DocTitle',
        parent=styles['Heading1'],
        fontSize=16,
        leading=20,
        alignment=1,
        textColor=colors.HexColor("#1A365D"),
        spaceAfter=6
    )
    sub_style = ParagraphStyle(
        'DocSub',
        parent=styles['Normal'],
        fontSize=10,
        leading=14,
        alignment=1,
        textColor=colors.HexColor("#4A5568")
    )
    cell_head = ParagraphStyle('CH', parent=styles['Normal'], fontSize=9, fontName='Helvetica-Bold', textColor=colors.whitesmoke, alignment=1)
    cell_body = ParagraphStyle('CB', parent=styles['Normal'], fontSize=8, leading=10)

    # Encabezado institucional
    story.append(Paragraph(f"<b>LISTA OFICIAL DE ESTUDIANTES</b>", title_style))
    story.append(Paragraph(f"<b>Diplomado:</b> {nombre_diplomado} &nbsp;|&nbsp; <b>Grupo / Clave:</b> {grupo_nombre}", sub_style))
    if instructores:
        story.append(Paragraph(f"<b>Instructor(es):</b> {instructores}", sub_style))
    story.append(Paragraph(f"<b>Fecha de Emisión:</b> {date.today().strftime('%d/%m/%Y')}", sub_style))
    story.append(Spacer(1, 15))

    # Construcción de la tabla (SOLO: #, Matrícula, Nombre, Carrera, Teléfono)
    data = [[
        Paragraph("#", cell_head),
        Paragraph("Matrícula", cell_head),
        Paragraph("Nombre Completo", cell_head),
        Paragraph("Carrera / Especialidad", cell_head),
        Paragraph("Teléfono", cell_head)
    ]]

    for idx, al in enumerate(st.session_state.alumnos, 1):
        data.append([
            Paragraph(str(idx), cell_body),
            Paragraph(al["matricula"], cell_body),
            Paragraph(al["nombre"], cell_body),
            Paragraph(al["carrera"], cell_body),
            Paragraph(al["telefono"], cell_body)
        ])

    col_widths = [25, 75, 200, 160, 80]
    t = Table(data, colWidths=col_widths, repeatRows=1)
    t.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#1A365D")),
        ('ALIGN', (0, 0), (0, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E0")),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor("#F7FAFC")]),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
        ('TOPPADDING', (0, 0), (-1, -1), 5),
    ]))
    story.append(t)
    doc.build(story)
    buffer.seek(0)
    return buffer

# --- INTERFAZ DE USUARIO ---
st.title("🎓 Sistema de Control Académico y Diplomados")

with st.sidebar:
    st.header("⚙️ Datos del Diplomado")
    nombre_diplomado = st.text_input("Nombre del Diplomado", value="Automatización e Instalaciones Especiales")
    grupo_clave = st.text_input("Grupo / Generación", value="Gpo. 2026-A")
    instructor = st.text_input("Docente a Cargo", value="Mtro. Omar Alejandro Jiménez Navarro")

    total_horas = sum(item["horas"] for item in st.session_state.bitacora_horas)
    st.metric("Total de Horas Impartidas", f"{total_horas:.1f} hrs")

    st.divider()
    st.subheader("📄 Emisión de Documentos")
    pdf_data = generar_reporte_pdf(nombre_diplomado, grupo_clave, instructor)
    st.download_button(
        label="📥 Descargar Reporte de Alumnos (PDF)",
        data=pdf_data,
        file_name=f"Reporte_Grupo_{grupo_clave.replace(' ', '_')}.pdf",
        mime="application/pdf",
        help="El reporte generado solo muestra Nombre, Matrícula, Carrera y Teléfono."
    )

# Pestañas principales
tab_alumnos, tab_asistencia, tab_horas = st.tabs([
    "📋 Alumnos y Control Financiero",
    "✅ Pase de Lista",
    "⏱️ Registro de Horas de Clase"
])

# --- TAB 1: GESTIÓN DE ALUMNOS Y PAGOS ---
with tab_alumnos:
    st.subheader("Listado de Estudiantes y Estado de Pagos")

    if st.session_state.alumnos:
        df_alumnos = pd.DataFrame(st.session_state.alumnos)
        df_alumnos.rename(columns={
            "matricula": "Matrícula",
            "nombre": "Nombre",
            "carrera": "Carrera",
            "telefono": "Teléfono",
            "pago_1": "1er Pago",
            "pago_2": "2do Pago",
            "pago_completo": "Pago Completo"
        }, inplace=True)

        col_config = {
            "Matrícula": st.column_config.TextColumn(required=True),
            "Nombre": st.column_config.TextColumn(required=True),
            "Carrera": st.column_config.TextColumn(),
            "Teléfono": st.column_config.TextColumn(),
            "1er Pago": st.column_config.CheckboxColumn(),
            "2do Pago": st.column_config.CheckboxColumn(),
            "Pago Completo": st.column_config.CheckboxColumn(),
        }

        # Editor de datos interactivo
        edited_df = st.data_editor(
            df_alumnos,
            use_container_width=True,
            num_rows="dynamic",
            column_config=col_config,
            key="editor_alumnos"
        )

        if st.button("💾 Guardar Cambios en la Lista"):
            st.session_state.alumnos = [
                {
                    "matricula": str(row["Matrícula"]),
                    "nombre": str(row["Nombre"]),
                    "carrera": str(row["Carrera"]),
                    "telefono": str(row["Teléfono"]),
                    "pago_1": bool(row["1er Pago"]),
                    "pago_2": bool(row["2do Pago"]),
                    "pago_completo": bool(row["Pago Completo"]),
                }
                for _, row in edited_df.iterrows()
            ]
            st.success("Cambios actualizados y guardados correctamente.")
            st.rerun()
    else:
        st.info("No hay alumnos registrados actualmente.")

    with st.expander("➕ Alta Rápida de Alumno"):
        with st.form("form_alta"):
            c1, c2 = st.columns(2)
            nueva_mat = c1.text_input("Matrícula")
            nuevo_nom = c2.text_input("Nombre Completo")
            c3, c4 = st.columns(2)
            nueva_car = c3.text_input("Carrera")
            nuevo_tel = c4.text_input("Teléfono")

            st.write("Estatus de Pago Inicial:")
            cp1, cp2, cp3 = st.columns(3)
            p1 = cp1.checkbox("1er Pago")
            p2 = cp2.checkbox("2do Pago")
            p_comp = cp3.checkbox("Pago Completo")

            if st.form_submit_button("Registrar Estudiante"):
                if nueva_mat and nuevo_nom:
                    st.session_state.alumnos.append({
                        "matricula": nueva_mat,
                        "nombre": nuevo_nom,
                        "carrera": nueva_car,
                        "telefono": nuevo_tel,
                        "pago_1": p1,
                        "pago_2": p2,
                        "pago_completo": p_comp
                    })
                    st.success(f"Alumno {nuevo_nom} registrado.")
                    st.rerun()
                else:
                    st.error("Matrícula y Nombre son obligatorios.")

# --- TAB 2: PASE DE LISTA ---
with tab_asistencia:
    st.subheader("Registro de Asistencia Diaria")
    col_f, _ = st.columns([1, 2])
    fecha_sesion = col_f.date_input("Fecha de la sesión", value=date.today())
    fecha_key = str(fecha_sesion)

    if fecha_key not in st.session_state.asistencias:
        st.session_state.asistencias[fecha_key] = {al["matricula"]: True for al in st.session_state.alumnos}

    if st.session_state.alumnos:
        with st.form("form_asistencia"):
            st.write(f"Sesión: **{fecha_key}**")
            estados = {}
            for al in st.session_state.alumnos:
                val_prev = st.session_state.asistencias[fecha_key].get(al["matricula"], True)
                c_nom, c_chk = st.columns([3, 1])
                c_nom.write(f"**{al['nombre']}** ({al['matricula']})")
                estados[al["matricula"]] = c_chk.checkbox("Presente", value=val_prev, key=f"att_{fecha_key}_{al['matricula']}")

            if st.form_submit_button("Guardar Pase de Lista"):
                st.session_state.asistencias[fecha_key] = estados
                st.success(f"Asistencia guardada para la fecha {fecha_key}.")
    else:
        st.info("Agregue alumnos en la primera pestaña para pasar lista.")

# --- TAB 3: CONTROL DE HORAS IMPARTIDAS ---
with tab_horas:
    st.subheader("Bitácora de Horas Docentes")
    col_reg1, col_reg2 = st.columns(2)

    with col_reg1:
        with st.form("form_horas"):
            st.write("Registrar Nueva Sesión")
            f_clase = st.date_input("Fecha", value=date.today(), key="fecha_clase")
            hrs = st.number_input("Horas impartidas", min_value=0.5, max_value=12.0, value=2.0, step=0.5)
            tema = st.text_input("Tema / Contenido cubierto")

            if st.form_submit_button("Registrar Horas"):
                st.session_state.bitacora_horas.append({
                    "fecha": str(f_clase),
                    "horas": hrs,
                    "tema": tema
                })
                st.success(f"Se registraron {hrs} horas de clase.")
                st.rerun()

    with col_reg2:
        st.write("Historial de Clases")
        if st.session_state.bitacora_horas:
            df_h = pd.DataFrame(st.session_state.bitacora_horas)
            df_h.rename(columns={"fecha": "Fecha", "horas": "Horas", "tema": "Tema"}, inplace=True)
            st.dataframe(df_h, use_container_width=True)
            if st.button("Limpiar Bitácora de Horas"):
                st.session_state.bitacora_horas = []
                st.rerun()
        else:
            st.caption("No hay registros de horas todavía.")