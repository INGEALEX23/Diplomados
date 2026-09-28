from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.drawing.image import Image as OpenpyxlImage

def generar_excel(grupo_id, con_pagos=False):
    wb = Workbook()
    ws = wb.active
    ws.title = "LISTA OFICIAL"
    ws.views.sheetView[0].showGridLines = True

    # Consultar datos del grupo y de los alumnos
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("SELECT diplomado, grupo, docente FROM grupos WHERE id = ?", (grupo_id,))
    res_grupo = c.fetchone()
    conn.close()

    nom_dip = res_grupo[0] if res_grupo else ""
    nom_gpo = res_grupo[1] if res_grupo else ""
    nom_doc = res_grupo[2] if res_grupo else ""

    df_al = get_alumnos_df(grupo_id)

    # ==========================
    # 1. ENCABEZADO INSTITUCIONAL
    # ==========================
    # Inserción de logotipos
    if os.path.exists(PATH_LOGO_UADEC):
        try:
            img_uadec = OpenpyxlImage(PATH_LOGO_UADEC)
            img_uadec.width = 75
            img_uadec.height = 75
            ws.add_image(img_uadec, "A1")
        except Exception:
            pass

    # Columna límite (E para sin pagos, H para con pagos)
    col_fin_letra = "H" if con_pagos else "E"
    col_fin_num = 8 if con_pagos else 5

    if os.path.exists(PATH_LOGO_POSGRADO):
        try:
            img_pos = OpenpyxlImage(PATH_LOGO_POSGRADO)
            img_pos.width = 75
            img_pos.height = 75
            ws.add_image(img_pos, f"{col_fin_letra}1")
        except Exception:
            pass

    # Título central
    ws.merge_cells(f"B1:{chr(64 + col_fin_num - 1)}1")
    ws["B1"] = "UNIVERSIDAD AUTÓNOMA DE COAHUILA"
    ws["B1"].font = Font(name="Calibri", size=13, bold=True, color="1A365D")
    ws["B1"].alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells(f"B2:{chr(64 + col_fin_num - 1)}2")
    ws["B2"] = "FACULTAD DE INGENIERÍA - COORDINACIÓN DE POSGRADO"
    ws["B2"].font = Font(name="Calibri", size=10, bold=True, color="4A5568")
    ws["B2"].alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells(f"B3:{chr(64 + col_fin_num - 1)}3")
    ws["B3"] = "REPORTE OFICIAL DE PARTICIPANTES" if not con_pagos else "CONTROL INTERNO Y REPORTE DE PAGOS"
    ws["B3"].font = Font(name="Calibri", size=9, bold=False, color="718096")
    ws["B3"].alignment = Alignment(horizontal="center", vertical="center")

    ws.row_dimensions[1].height = 24
    ws.row_dimensions[2].height = 18
    ws.row_dimensions[3].height = 16
    ws.row_dimensions[4].height = 10

    # ==========================
    # 2. TARJETA DE DATOS (METADATOS)
    # ==========================
    fill_meta = PatternFill(start_color="EDF2F7", end_color="EDF2F7", fill_type="solid")
    font_meta = Font(name="Calibri", size=9, color="1A202C")
    thin_border_side = Side(style='thin', color="CBD5E0")
    border_meta = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)

    # Fila 5: Diplomado y Grupo
    ws["A5"] = f"DIPLOMADO: {nom_dip}"
    ws["A5"].font = font_meta
    ws["D5"] = f"GRUPO / GENERACIÓN: {nom_gpo}"
    ws["D5"].font = font_meta

    # Fila 6: Docente y Fecha
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

    ws.row_dimensions[7].height = 12  # Espaciador

    # ==========================
    # 3. TABLA DE ESTUDIANTES
    # ==========================
    if not con_pagos:
        headers = ["#", "MATRÍCULA", "NOMBRE COMPLETO", "CARRERA / ESPECIALIDAD", "TELÉFONO"]
    else:
        headers = ["#", "MATRÍCULA", "NOMBRE COMPLETO", "CARRERA / ESPECIALIDAD", "TELÉFONO", "1ER PAGO", "2DO PAGO", "PAGO COMPLETO"]

    # Fila de Encabezados (Azul Marino Institucional)
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

    # Filas de Alumnos con Zebra Striping (Blanco y Gris muy tenue)
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

        valores = [
            idx + 1,
            row["matricula"],
            row["nombre"],
            row["carrera"],
            row["telefono"]
        ]

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

            # Alineación (# y Matrícula centrados, pagos centrados, textos a la izquierda)
            if c_idx in [1, 2] or (con_pagos and c_idx >= 6):
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)

        fila_actual += 1

    # ==========================
    # 4. ANCHOS DE COLUMNAS
    # ==========================
    anchos = {
        "A": 6,    # #
        "B": 15,   # Matrícula
        "C": 38,   # Nombre
        "D": 32,   # Carrera
        "E": 16,   # Teléfono
    }
    if con_pagos:
        anchos["F"] = 14  # 1er Pago
        anchos["G"] = 14  # 2do Pago
        anchos["H"] = 16  # Pago Completo

    for col_letter, width in anchos.items():
        ws.column_dimensions[col_letter].width = width

    # Si es el archivo con pagos, agregar la segunda hoja con la Bitácora de Horas
    if con_pagos:
        df_h = get_horas_df(grupo_id)
        if not df_h.empty:
            ws_h = wb.create_sheet(title="BITÁCORA DE HORAS")
            ws_h.views.sheetView[0].showGridLines = True

            # Encabezado bitácora
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
