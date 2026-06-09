import pandas as pd

# 1) Leer archivo original (NUEVO NOMBRE)
archivo_origen = "Base_Produccion_Electronica.xlsx"
df = pd.read_excel(archivo_origen)

# 2) Corrección: convertir eficiencia a proporción
df["Eficiencia (%)"] = df["Eficiencia (%)"] / 100

# 3) Calcular columnas adicionales
df["Cumplimiento_%"] = df["Producción Real"] / df["Producción Plan"]
df["Productividad_x_Operador"] = df["Producción Real"] / df["Operadores"]

# 4) Crear archivo nuevo con xlsxwriter
salida = "Dashboard_Produccion.xlsx"
with pd.ExcelWriter(salida, engine="xlsxwriter") as writer:
    hoja_datos = "Datos"
    df.to_excel(writer, sheet_name=hoja_datos, index=False)
    workbook = writer.book
    ws_datos = writer.sheets[hoja_datos]

    # Formatos
    formato_porcentaje = workbook.add_format({"num_format": "0.0%", "font_name": "Segoe UI"})
    formato_numero = workbook.add_format({"num_format": "#,##0", "font_name": "Segoe UI"})
    formato_texto = workbook.add_format({"font_name": "Segoe UI"})

    # Ajuste de columnas
    ws_datos.set_column("A:A", 12, formato_texto)
    ws_datos.set_column("B:C", 14, formato_numero)
    ws_datos.set_column("D:D", 10, formato_porcentaje)
    ws_datos.set_column("E:E", 14, formato_numero)
    ws_datos.set_column("F:F", 12, formato_numero)
    ws_datos.set_column("G:G", 12, formato_porcentaje)
    ws_datos.set_column("H:H", 14, formato_porcentaje)
    ws_datos.set_column("I:I", 18, formato_numero)

    # --- Hoja Cálculos ---
    hoja_calc = "Cálculos"
    ws_calc = workbook.add_worksheet(hoja_calc)

    formato_titulo = workbook.add_format({"bold": True, "font_size": 12, "font_name": "Segoe UI"})
    formato_kpi = workbook.add_format({"bold": True, "font_size": 14, "font_name": "Segoe UI"})

    ws_calc.write("B2", "KPIs Anuales", formato_titulo)
    ws_calc.write("B4", "Cumplimiento promedio anual", formato_texto)
    ws_calc.write("B5", "Productividad promedio anual", formato_texto)
    ws_calc.write("B6", "Scrap promedio anual", formato_texto)
    ws_calc.write("B7", "Eficiencia promedio anual", formato_texto)
    ws_calc.write("B8", "Tiempo muerto total anual", formato_texto)
    ws_calc.write("B9", "Producción total anual", formato_texto)

    ws_calc.write_formula("C4", "=PROMEDIO(Tabla_Produccion[Cumplimiento_%])", formato_kpi)
    ws_calc.write_formula("C5", "=PROMEDIO(Tabla_Produccion[Productividad_x_Operador])", formato_kpi)
    ws_calc.write_formula("C6", "=PROMEDIO(Tabla_Produccion[Scrap (%)])", formato_kpi)
    ws_calc.write_formula("C7", "=PROMEDIO(Tabla_Produccion[Eficiencia (%)])", formato_kpi)
    ws_calc.write_formula("C8", "=SUMA(Tabla_Produccion[Tiempo Muerto (h)])", formato_kpi)
    ws_calc.write_formula("C9", "=SUMA(Tabla_Produccion[Producción Real])", formato_kpi)

    # --- Hoja Dashboard ---
    hoja_dash = "Dashboard"
    ws_dash = workbook.add_worksheet(hoja_dash)

    gris_claro = workbook.add_format({"bg_color": "#F2F2F2", "font_name": "Segoe UI"})
    formato_kpi_card = workbook.add_format({"bg_color": "#F2F2F2", "font_name": "Segoe UI"})
    formato_kpi_titulo = workbook.add_format({"font_name": "Segoe UI", "bold": True, "align": "center"})
    formato_kpi_valor = workbook.add_format({"font_name": "Segoe UI", "bold": True, "font_size": 18, "align": "center"})

    ws_dash.merge_range("B1:K1", "DASHBOARD DE PRODUCCIÓN",
                        workbook.add_format({"font_name": "Segoe UI", "bold": True, "font_size": 16, "align": "center"}))

    # Tarjetas KPI
    kpis = [
        ("Cumplimiento (%)", "C4"),
        ("Eficiencia (%)", "C7"),
        ("Scrap promedio (%)", "C6"),
        ("Productividad x Oper.", "C5"),
        ("Producción total", "C9"),
    ]

    tarjetas_cols = [("B", "C"), ("D", "E"), ("F", "G"), ("H", "I"), ("J", "K")]

    for (titulo, celda), (col_ini, col_fin) in zip(kpis, tarjetas_cols):
        col_start = ord(col_ini) - ord("A")
        col_end = ord(col_fin) - ord("A")

        # Fondo
        for row in range(3, 8):
            for col in range(col_start, col_end + 1):
                ws_dash.write_blank(row - 1, col, None, formato_kpi_card)

        ws_dash.merge_range(f"{col_ini}3:{col_fin}3", titulo, formato_kpi_titulo)
        ws_dash.merge_range(f"{col_ini}4:{col_fin}6", f"='{hoja_calc}'!{celda}", formato_kpi_valor)

    # --- Gráficos ---
    n = len(df)
    fila_ini = 2
    fila_fin = fila_ini + n - 1

    def rango(col, fi, ff):
        return f"='{hoja_datos}'!{col}{fi}:{col}{ff}"

    # Cumplimiento
    chart_cump = workbook.add_chart({"type": "line"})
    chart_cump.add_series({
        "name": "Cumplimiento (%)",
        "categories": rango("A", fila_ini, fila_fin),
        "values": rango("H", fila_ini, fila_fin),
        "line": {"color": "#1976D2", "width": 2},
    })
    ws_dash.insert_chart("B9", chart_cump)

    # Eficiencia
    chart_ef = workbook.add_chart({"type": "line"})
    chart_ef.add_series({
        "name": "Eficiencia (%)",
        "categories": rango("A", fila_ini, fila_fin),
        "values": rango("G", fila_ini, fila_fin),
        "line": {"color": "#388E3C", "width": 2},
    })
    ws_dash.insert_chart("G9", chart_ef)

    # Scrap
    chart_scrap = workbook.add_chart({"type": "column"})
    chart_scrap.add_series({
        "name": "Scrap (%)",
        "categories": rango("A", fila_ini, fila_fin),
        "values": rango("D", fila_ini, fila_fin),
        "fill": {"color": "#64B5F6"},
    })
    ws_dash.insert_chart("B22", chart_scrap)

    # Productividad
    chart_prod = workbook.add_chart({"type": "column"})
    chart_prod.add_series({
        "name": "Productividad x Operador",
        "categories": rango("A", fila_ini, fila_fin),
        "values": rango("I", fila_ini, fila_fin),
        "fill": {"color": "#90CAF9"},
    })
    ws_dash.insert_chart("G22", chart_prod)

    # Dispersión
    chart_disp = workbook.add_chart({"type": "scatter"})
    chart_disp.add_series({
        "name": "Tiempo Muerto vs Eficiencia",
        "categories": rango("E", fila_ini, fila_fin),
        "values": rango("G", fila_ini, fila_fin),
        "marker": {"type": "circle", "size": 6, "border": {"color": "#1976D2"}},
    })
    ws_dash.insert_chart("B36", chart_disp)

print("Dashboard_Produccion.xlsx generado correctamente.")
