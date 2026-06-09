import pandas as pd
import numpy as np
from datetime import datetime, timedelta

# =========================================================
# Generación de datos simulados (30 días, 2 turnos)
# =========================================================

def generar_datos():
    start_date = datetime(2026, 5, 1)
    dates = [start_date + timedelta(days=i) for i in range(30)]

    rows = []
    for d in dates:
        for turno in [1, 2]:
            # Línea alternada por día (para tener L1 y L2 en el dataset)
            linea = "L1" if d.day % 2 == 1 else "L2"
            produccion_plan = 2100
            produccion_real = produccion_plan - np.random.randint(0, 121)
            scrap = round(np.random.uniform(0.020, 0.027), 3)          # decimal
            tiempo_muerto = round(np.random.uniform(0.5, 1.3), 1)
            operadores = np.random.randint(25, 28)
            eficiencia = round(np.random.uniform(0.95, 0.99), 3)       # decimal

            rows.append([
                d,
                turno,
                linea,
                produccion_plan,
                produccion_real,
                scrap,
                tiempo_muerto,
                operadores,
                eficiencia
            ])

    df = pd.DataFrame(
        rows,
        columns=[
            "Fecha",
            "Turno",
            "Línea",
            "Producción Plan",
            "Producción Real",
            "Scrap (%)",
            "Tiempo Muerto (h)",
            "Operadores",
            "Eficiencia (%)"
        ]
    )

    # Campos de calendario
    df["Fecha"] = pd.to_datetime(df["Fecha"])
    dias_es = {
        0: "Lunes",
        1: "Martes",
        2: "Miércoles",
        3: "Jueves",
        4: "Viernes",
        5: "Sábado",
        6: "Domingo",
    }
    df["Día"] = df["Fecha"].dt.dayofweek.map(dias_es)
    df["Semana"] = df["Fecha"].dt.isocalendar().week.astype(int)
    meses_es = {
        1: "enero",
        2: "febrero",
        3: "marzo",
        4: "abril",
        5: "mayo",
        6: "junio",
        7: "julio",
        8: "agosto",
        9: "septiembre",
        10: "octubre",
        11: "noviembre",
        12: "diciembre",
    }
    df["Mes"] = df["Fecha"].dt.month.map(meses_es)

    # Reordenar columnas
    df = df[
        [
            "Fecha",
            "Día",
            "Semana",
            "Mes",
            "Turno",
            "Línea",
            "Producción Plan",
            "Producción Real",
            "Scrap (%)",
            "Tiempo Muerto (h)",
            "Operadores",
            "Eficiencia (%)",
        ]
    ]

    # Campos calculados
    df["Cumplimiento (%)"] = df["Producción Real"] / df["Producción Plan"]
    df["Productividad por Operador"] = df["Producción Real"] / df["Operadores"]

    df = df.sort_values("Fecha").reset_index(drop=True)

    df["Producción Acumulada"] = df["Producción Real"].cumsum()
    df["Unidades Scrap"] = df["Producción Real"] * df["Scrap (%)"]
    df["Scrap Acumulado (%)"] = (
        df["Unidades Scrap"].cumsum() / df["Producción Real"].cumsum()
    )

    return df


# =========================================================
# Creación del Excel V2 con diseño ejecutivo
# =========================================================

def crear_excel_v2(df, nombre_archivo="Dashboard_Produccion_Diario_Semanal_Mensual_V2.xlsx"):
    with pd.ExcelWriter(nombre_archivo, engine="xlsxwriter", datetime_format="dd/mm/yyyy") as writer:
        workbook = writer.book

        # -------------------------
        # Hoja Datos_Diarios
        # -------------------------
        df_export = df.drop(columns=["Unidades Scrap"])
        df_export.to_excel(writer, sheet_name="Datos_Diarios", index=False)
        ws_datos = writer.sheets["Datos_Diarios"]

        # Formatos base
        fmt_header = workbook.add_format(
            {"bold": True, "bg_color": "#D9E1F2", "border": 1}
        )
        fmt_date = workbook.add_format({"num_format": "dd/mm/yyyy"})
        fmt_num = workbook.add_format({"num_format": "#,##0"})
        fmt_pct = workbook.add_format({"num_format": "0.00%"})
        fmt_text = workbook.add_format({})
        fmt_gray = workbook.add_format({"bg_color": "#F2F2F2"})
        fmt_kpi_title = workbook.add_format(
            {
                "bold": True,
                "font_size": 10,
                "bg_color": "#E6E6E6",
                "border": 1,
                "align": "center",
                "valign": "vcenter",
            }
        )
        fmt_kpi_value = workbook.add_format(
            {
                "bold": True,
                "font_size": 14,
                "border": 1,
                "align": "right",
                "valign": "vcenter",
            }
        )
        fmt_kpi_value_pct = workbook.add_format(
            {
                "bold": True,
                "font_size": 14,
                "border": 1,
                "align": "right",
                "valign": "vcenter",
                "num_format": "0.00%",
            }
        )

        fmt_green = workbook.add_format({"bg_color": "#C6EFCE", "font_color": "#006100"})
        fmt_yellow = workbook.add_format({"bg_color": "#FFEB9C", "font_color": "#9C6500"})
        fmt_red = workbook.add_format({"bg_color": "#F8CBAD", "font_color": "#9C0006"})

        # Encabezados
        ws_datos.set_row(0, None, fmt_header)

        col_indices = {col: i for i, col in enumerate(df_export.columns)}

        # Formatos por columna
        ws_datos.set_column(col_indices["Fecha"], col_indices["Fecha"], 12, fmt_date)
        ws_datos.set_column(col_indices["Día"], col_indices["Día"], 12, fmt_text)
        ws_datos.set_column(col_indices["Mes"], col_indices["Mes"], 12, fmt_text)
        ws_datos.set_column(col_indices["Semana"], col_indices["Semana"], 8, fmt_num)
        ws_datos.set_column(col_indices["Turno"], col_indices["Turno"], 8, fmt_num)
        ws_datos.set_column(col_indices["Línea"], col_indices["Línea"], 8, fmt_text)

        for col in ["Producción Plan", "Producción Real", "Tiempo Muerto (h)",
                    "Operadores", "Productividad por Operador", "Producción Acumulada"]:
            ws_datos.set_column(col_indices[col], col_indices[col], 16, fmt_num)

        for col in ["Scrap (%)", "Eficiencia (%)", "Cumplimiento (%)", "Scrap Acumulado (%)"]:
            ws_datos.set_column(col_indices[col], col_indices[col], 16, fmt_pct)

        # -------------------------
        # Hoja Resumen
        # -------------------------
        ws_resumen = workbook.add_worksheet("Resumen")

        # Resumen diario
        res_dia = (
            df.groupby("Fecha")
            .agg(
                {
                    "Producción Real": "sum",
                    "Producción Plan": "sum",
                    "Tiempo Muerto (h)": "sum",
                    "Scrap (%)": "mean",
                    "Eficiencia (%)": "mean",
                    "Productividad por Operador": "mean",
                }
            )
            .reset_index()
        )
        res_dia["Cumplimiento (%)"] = res_dia["Producción Real"] / res_dia["Producción Plan"]

        res_dia_cols = [
            "Fecha",
            "Producción Real",
            "Producción Plan",
            "Cumplimiento (%)",
            "Scrap (%)",
            "Eficiencia (%)",
            "Tiempo Muerto (h)",
            "Productividad por Operador",
        ]
        res_dia = res_dia[res_dia_cols]

        ws_resumen.write("A1", "Resumen Diario", fmt_header)
        for j, col in enumerate(res_dia_cols):
            ws_resumen.write(1, j, col, fmt_header)
        for i, row in res_dia.iterrows():
            ws_resumen.write_datetime(i + 2, 0, row["Fecha"], fmt_date)
            ws_resumen.write_number(i + 2, 1, row["Producción Real"], fmt_num)
            ws_resumen.write_number(i + 2, 2, row["Producción Plan"], fmt_num)
            ws_resumen.write_number(i + 2, 3, row["Cumplimiento (%)"], fmt_pct)
            ws_resumen.write_number(i + 2, 4, row["Scrap (%)"], fmt_pct)
            ws_resumen.write_number(i + 2, 5, row["Eficiencia (%)"], fmt_pct)
            ws_resumen.write_number(i + 2, 6, row["Tiempo Muerto (h)"], fmt_num)
            ws_resumen.write_number(i + 2, 7, row["Productividad por Operador"], fmt_num)

        n_dias = len(res_dia)

        # Resumen semanal
        start_row_sem = n_dias + 4
        res_sem = (
            df.groupby("Semana")
            .agg(
                {
                    "Producción Real": "sum",
                    "Producción Plan": "sum",
                    "Tiempo Muerto (h)": "sum",
                    "Scrap (%)": "mean",
                    "Eficiencia (%)": "mean",
                    "Productividad por Operador": "mean",
                }
            )
            .reset_index()
        )
        res_sem["Cumplimiento (%)"] = res_sem["Producción Real"] / res_sem["Producción Plan"]

        res_sem_cols = [
            "Semana",
            "Producción Real",
            "Producción Plan",
            "Cumplimiento (%)",
            "Scrap (%)",
            "Eficiencia (%)",
            "Tiempo Muerto (h)",
            "Productividad por Operador",
        ]

        ws_resumen.write(start_row_sem, 0, "Resumen Semanal", fmt_header)
        for j, col in enumerate(res_sem_cols):
            ws_resumen.write(start_row_sem + 1, j, col, fmt_header)
        for i, row in res_sem.iterrows():
            ws_resumen.write_number(start_row_sem + 2 + i, 0, row["Semana"], fmt_num)
            ws_resumen.write_number(start_row_sem + 2 + i, 1, row["Producción Real"], fmt_num)
            ws_resumen.write_number(start_row_sem + 2 + i, 2, row["Producción Plan"], fmt_num)
            ws_resumen.write_number(start_row_sem + 2 + i, 3, row["Cumplimiento (%)"], fmt_pct)
            ws_resumen.write_number(start_row_sem + 2 + i, 4, row["Scrap (%)"], fmt_pct)
            ws_resumen.write_number(start_row_sem + 2 + i, 5, row["Eficiencia (%)"], fmt_pct)
            ws_resumen.write_number(start_row_sem + 2 + i, 6, row["Tiempo Muerto (h)"], fmt_num)
            ws_resumen.write_number(start_row_sem + 2 + i, 7, row["Productividad por Operador"], fmt_num)

        # Resumen mensual
        start_row_mes = start_row_sem + len(res_sem) + 5
        res_mes = (
            df.groupby("Mes")
            .agg(
                {
                    "Producción Real": "sum",
                    "Producción Plan": "sum",
                    "Tiempo Muerto (h)": "sum",
                    "Scrap (%)": "mean",
                    "Eficiencia (%)": "mean",
                    "Productividad por Operador": "mean",
                }
            )
            .reset_index()
        )
        res_mes["Cumplimiento (%)"] = res_mes["Producción Real"] / res_mes["Producción Plan"]

        res_mes_cols = [
            "Mes",
            "Producción Real",
            "Producción Plan",
            "Cumplimiento (%)",
            "Scrap (%)",
            "Eficiencia (%)",
            "Tiempo Muerto (h)",
            "Productividad por Operador",
        ]

        ws_resumen.write(start_row_mes, 0, "Resumen Mensual", fmt_header)
        for j, col in enumerate(res_mes_cols):
            ws_resumen.write(start_row_mes + 1, j, col, fmt_header)
        for i, row in res_mes.iterrows():
            ws_resumen.write(start_row_mes + 2 + i, 0, row["Mes"])
            ws_resumen.write_number(start_row_mes + 2 + i, 1, row["Producción Real"], fmt_num)
            ws_resumen.write_number(start_row_mes + 2 + i, 2, row["Producción Plan"], fmt_num)
            ws_resumen.write_number(start_row_mes + 2 + i, 3, row["Cumplimiento (%)"], fmt_pct)
            ws_resumen.write_number(start_row_mes + 2 + i, 4, row["Scrap (%)"], fmt_pct)
            ws_resumen.write_number(start_row_mes + 2 + i, 5, row["Eficiencia (%)"], fmt_pct)
            ws_resumen.write_number(start_row_mes + 2 + i, 6, row["Tiempo Muerto (h)"], fmt_num)
            ws_resumen.write_number(start_row_mes + 2 + i, 7, row["Productividad por Operador"], fmt_num)

        # Resumen por turno
        res_turno = (
            df.groupby("Turno")
            .agg(
                {
                    "Producción Real": "sum",
                    "Producción Plan": "sum",
                    "Scrap (%)": "mean",
                    "Eficiencia (%)": "mean",
                }
            )
            .reset_index()
        )
        res_turno["Cumplimiento (%)"] = res_turno["Producción Real"] / res_turno["Producción Plan"]

        # Resumen por línea
        res_linea = (
            df.groupby("Línea")
            .agg(
                {
                    "Producción Real": "sum",
                    "Producción Plan": "sum",
                    "Scrap (%)": "mean",
                    "Eficiencia (%)": "mean",
                    "Tiempo Muerto (h)": "sum",
                }
            )
            .reset_index()
        )
        res_linea["Cumplimiento (%)"] = res_linea["Producción Real"] / res_linea["Producción Plan"]

        # -------------------------
        # Hoja Dashboard
        # -------------------------
        ws_dash = workbook.add_worksheet("Dashboard")
        ws_dash.hide_gridlines(2)  # quitar líneas de cuadrícula

        ws_dash.set_column("A:A", 2)   # margen
        ws_dash.set_column("B:M", 16)

        # Título
        title_fmt = workbook.add_format(
            {"bold": True, "font_size": 16, "align": "left", "valign": "vcenter"}
        )
        ws_dash.write("B1", "Dashboard de Producción Diario / Semanal / Mensual", title_fmt)

        # -------------------------
        # Tarjetas KPI superiores
        # -------------------------
        # Producción Total
        ws_dash.write("B3", "Producción Total", fmt_kpi_title)
        ws_dash.write_formula("B4", "=SUM(Datos_Diarios!H:H)", fmt_kpi_value)

        # Cumplimiento del Plan
        ws_dash.write("D3", "Cumplimiento del Plan", fmt_kpi_title)
        ws_dash.write_formula(
            "D4",
            "=IF(SUM(Datos_Diarios!G:G)>0,SUM(Datos_Diarios!H:H)/SUM(Datos_Diarios!G:G),0)",
            fmt_kpi_value_pct,
        )

        # Eficiencia
        ws_dash.write("F3", "Eficiencia", fmt_kpi_title)
        ws_dash.write_formula(
            "F4",
            "=AVERAGE(Datos_Diarios!L:L)",
            fmt_kpi_value_pct,
        )

        # Scrap
        ws_dash.write("H3", "Scrap", fmt_kpi_title)
        ws_dash.write_formula(
            "H4",
            "=AVERAGE(Datos_Diarios!I:I)",
            fmt_kpi_value_pct,
        )

        # Productividad por Operador
        ws_dash.write("J3", "Productividad por Operador", fmt_kpi_title)
        ws_dash.write_formula(
            "J4",
            "=AVERAGE(Datos_Diarios!N:N)",
            fmt_kpi_value,
        )

        # Tiempo Muerto
        ws_dash.write("L3", "Tiempo Muerto (h)", fmt_kpi_title)
        ws_dash.write_formula(
            "L4",
            "=SUM(Datos_Diarios!J:J)",
            fmt_kpi_value,
        )

        # -------------------------
        # Semáforos (condicional)
        # -------------------------
        # Cumplimiento (D4)
        ws_dash.conditional_format(
            "D4",
            {
                "type": "formula",
                "criteria": "=D4>=0.97",
                "format": fmt_green,
            },
        )
        ws_dash.conditional_format(
            "D4",
            {
                "type": "formula",
                "criteria": "=AND(D4>=0.95,D4<0.97)",
                "format": fmt_yellow,
            },
        )
        ws_dash.conditional_format(
            "D4",
            {
                "type": "formula",
                "criteria": "=D4<0.95",
                "format": fmt_red,
            },
        )

        # Eficiencia (F4)
        ws_dash.conditional_format(
            "F4",
            {
                "type": "formula",
                "criteria": "=F4>=0.97",
                "format": fmt_green,
            },
        )
        ws_dash.conditional_format(
            "F4",
            {
                "type": "formula",
                "criteria": "=AND(F4>=0.95,F4<0.97)",
                "format": fmt_yellow,
            },
        )
        ws_dash.conditional_format(
            "F4",
            {
                "type": "formula",
                "criteria": "=F4<0.95",
                "format": fmt_red,
            },
        )

        # Scrap (H4) – menor es mejor
        ws_dash.conditional_format(
            "H4",
            {
                "type": "formula",
                "criteria": "=H4<=0.023",
                "format": fmt_green,
            },
        )
        ws_dash.conditional_format(
            "H4",
            {
                "type": "formula",
                "criteria": "=AND(H4>0.023,H4<=0.026)",
                "format": fmt_yellow,
            },
        )
        ws_dash.conditional_format(
            "H4",
            {
                "type": "formula",
                "criteria": "=H4>0.026",
                "format": fmt_red,
            },
        )

        # -------------------------
        # Comparativo por Turno
        # -------------------------
        start_row_turno = 7
        ws_dash.write(start_row_turno, 2, "Comparativo por Turno", fmt_header)
        turno_cols = ["Turno", "Producción Real", "Cumplimiento (%)", "Eficiencia (%)", "Scrap (%)"]
        for j, col in enumerate(turno_cols):
            ws_dash.write(start_row_turno + 1, 2 + j, col, fmt_header)

        for i, row in res_turno.iterrows():
            r = start_row_turno + 2 + i
            ws_dash.write_number(r, 2, row["Turno"], fmt_num)
            ws_dash.write_number(r, 3, row["Producción Real"], fmt_num)
            ws_dash.write_number(r, 4, row["Cumplimiento (%)"], fmt_pct)
            ws_dash.write_number(r, 5, row["Eficiencia (%)"], fmt_pct)
            ws_dash.write_number(r, 6, row["Scrap (%)"], fmt_pct)

        # -------------------------
        # Comparativo por Línea
        # -------------------------
        start_row_linea = start_row_turno + 6
        ws_dash.write(start_row_linea, 2, "Comparativo por Línea", fmt_header)
        linea_cols = ["Línea", "Producción Real", "Cumplimiento (%)", "Eficiencia (%)", "Scrap (%)", "Tiempo Muerto (h)"]
        for j, col in enumerate(linea_cols):
            ws_dash.write(start_row_linea + 1, 2 + j, col, fmt_header)

        for i, row in res_linea.iterrows():
            r = start_row_linea + 2 + i
            ws_dash.write(r, 2, row["Línea"])
            ws_dash.write_number(r, 3, row["Producción Real"], fmt_num)
            ws_dash.write_number(r, 4, row["Cumplimiento (%)"], fmt_pct)
            ws_dash.write_number(r, 5, row["Eficiencia (%)"], fmt_pct)
            ws_dash.write_number(r, 6, row["Scrap (%)"], fmt_pct)
            ws_dash.write_number(r, 7, row["Tiempo Muerto (h)"], fmt_num)

        # -------------------------
        # Insights Ejecutivos
        # -------------------------
        insight_title_fmt = workbook.add_format(
            {"bold": True, "font_size": 11, "bg_color": "#D9E1F2", "border": 1}
        )
        insight_text_fmt = workbook.add_format(
            {"font_size": 10, "border": 1}
        )

        ws_dash.write("J7", "Insights Ejecutivos", insight_title_fmt)

        # Mejor semana por cumplimiento
        best_week = res_sem.sort_values("Cumplimiento (%)", ascending=False).iloc[0]
        ws_dash.write("J8", "Mejor semana por cumplimiento:", insight_text_fmt)
        ws_dash.write("K8", f"Semana {int(best_week['Semana'])}", insight_text_fmt)

        # Día con mayor producción
        best_day = res_dia.sort_values("Producción Real", ascending=False).iloc[0]
        ws_dash.write("J9", "Día con mayor producción:", insight_text_fmt)
        ws_dash.write_datetime("K9", best_day["Fecha"], fmt_date)

        # Día con mayor scrap
        day_max_scrap = df.sort_values("Scrap (%)", ascending=False).iloc[0]
        ws_dash.write("J10", "Día con mayor scrap:", insight_text_fmt)
        ws_dash.write_datetime("K10", day_max_scrap["Fecha"], fmt_date)

        # Día con menor eficiencia
        day_min_eff = df.sort_values("Eficiencia (%)", ascending=True).iloc[0]
        ws_dash.write("J11", "Día con menor eficiencia:", insight_text_fmt)
        ws_dash.write_datetime("K11", day_min_eff["Fecha"], fmt_date)

        # Línea con mayor tiempo muerto
        line_max_tm = res_linea.sort_values("Tiempo Muerto (h)", ascending=False).iloc[0]
        ws_dash.write("J12", "Línea con mayor tiempo muerto:", insight_text_fmt)
        ws_dash.write("K12", line_max_tm["Línea"], insight_text_fmt)

        # Turno con mejor cumplimiento
        turno_best = res_turno.sort_values("Cumplimiento (%)", ascending=False).iloc[0]
        ws_dash.write("J13", "Turno con mejor cumplimiento:", insight_text_fmt)
        ws_dash.write_number("K13", turno_best["Turno"], insight_text_fmt)

        # -------------------------
        # Gráficos (alineados, colores sobrios)
        # -------------------------
        # Producción diaria
        chart_prod_dia = workbook.add_chart({"type": "column"})
        chart_prod_dia.add_series(
            {
                "name": "Producción Real diaria",
                "categories": ["Resumen", 2, 0, 1 + n_dias, 0],
                "values": ["Resumen", 2, 1, 1 + n_dias, 1],
                "fill": {"color": "#4472C4"},
                "border": {"color": "#4472C4"},
            }
        )
        chart_prod_dia.set_title({"name": "Producción diaria"})
        chart_prod_dia.set_x_axis({"name": "Fecha"})
        chart_prod_dia.set_y_axis({"name": "Piezas"})
        chart_prod_dia.set_style(2)
        ws_dash.insert_chart("B18", chart_prod_dia, {"x_scale": 1.4, "y_scale": 1.2})

        # Producción semanal
        n_sem = len(res_sem)
        chart_prod_sem = workbook.add_chart({"type": "column"})
        chart_prod_sem.add_series(
            {
                "name": "Producción Real semanal",
                "categories": ["Resumen", start_row_sem + 2, 0, start_row_sem + 1 + n_sem, 0],
                "values": ["Resumen", start_row_sem + 2, 1, start_row_sem + 1 + n_sem, 1],
                "fill": {"color": "#5B9BD5"},
                "border": {"color": "#5B9BD5"},
            }
        )
        chart_prod_sem.set_title({"name": "Producción semanal"})
        chart_prod_sem.set_x_axis({"name": "Semana"})
        chart_prod_sem.set_y_axis({"name": "Piezas"})
        chart_prod_sem.set_style(2)
        ws_dash.insert_chart("H18", chart_prod_sem, {"x_scale": 1.4, "y_scale": 1.2})

        # Producción mensual
        n_mes = len(res_mes)
        chart_prod_mes = workbook.add_chart({"type": "column"})
        chart_prod_mes.add_series(
            {
                "name": "Producción Real mensual",
                "categories": ["Resumen", start_row_mes + 2, 0, start_row_mes + 1 + n_mes, 0],
                "values": ["Resumen", start_row_mes + 2, 1, start_row_mes + 1 + n_mes, 1],
                "fill": {"color": "#A5A5A5"},
                "border": {"color": "#A5A5A5"},
            }
        )
        chart_prod_mes.set_title({"name": "Producción mensual"})
        chart_prod_mes.set_x_axis({"name": "Mes"})
        chart_prod_mes.set_y_axis({"name": "Piezas"})
        chart_prod_mes.set_style(2)
        ws_dash.insert_chart("B34", chart_prod_mes, {"x_scale": 1.4, "y_scale": 1.2})

        # Tendencia de eficiencia
        chart_eff = workbook.add_chart({"type": "line"})
        chart_eff.add_series(
            {
                "name": "Eficiencia diaria",
                "categories": ["Resumen", 2, 0, 1 + n_dias, 0],
                "values": ["Resumen", 2, 5, 1 + n_dias, 5],
                "line": {"color": "#4472C4"},
            }
        )
        chart_eff.set_title({"name": "Tendencia de eficiencia"})
        chart_eff.set_y_axis({"num_format": "0.00%"})
        chart_eff.set_style(2)
        ws_dash.insert_chart("H34", chart_eff, {"x_scale": 1.4, "y_scale": 1.2})

        # Tendencia de scrap
        chart_scrap = workbook.add_chart({"type": "line"})
        chart_scrap.add_series(
            {
                "name": "Scrap diario",
                "categories": ["Resumen", 2, 0, 1 + n_dias, 0],
                "values": ["Resumen", 2, 4, 1 + n_dias, 4],
                "line": {"color": "#A5A5A5"},
            }
        )
        chart_scrap.set_title({"name": "Tendencia de scrap"})
        chart_scrap.set_y_axis({"num_format": "0.00%"})
        chart_scrap.set_style(2)
        ws_dash.insert_chart("B50", chart_scrap, {"x_scale": 1.4, "y_scale": 1.2})

        # Tiempo muerto vs eficiencia
        chart_tm_eff = workbook.add_chart({"type": "scatter"})
        chart_tm_eff.add_series(
            {
                "name": "Tiempo muerto vs eficiencia",
                "categories": ["Resumen", 2, 6, 1 + n_dias, 6],  # Tiempo Muerto (h)
                "values": ["Resumen", 2, 5, 1 + n_dias, 5],      # Eficiencia (%)
                "marker": {"type": "circle", "size": 5, "border": {"color": "#4472C4"}, "fill": {"color": "#4472C4"}},
            }
        )
        chart_tm_eff.set_title({"name": "Tiempo muerto vs eficiencia"})
        chart_tm_eff.set_x_axis({"name": "Tiempo Muerto (h)"})
        chart_tm_eff.set_y_axis({"name": "Eficiencia (%)", "num_format": "0.00%"})
        chart_tm_eff.set_style(2)
        ws_dash.insert_chart("H50", chart_tm_eff, {"x_scale": 1.4, "y_scale": 1.2})

    print(f"Archivo generado: {nombre_archivo}")


# =========================================================
# Punto de entrada
# =========================================================

if __name__ == "__main__":
    df_datos = generar_datos()
    crear_excel_v2(df_datos)
