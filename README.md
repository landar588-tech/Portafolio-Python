# Dashboard de Producción Diario, Semanal y Mensual

## Descripción

Proyecto desarrollado para monitorear indicadores clave de manufactura electrónica utilizando Python y Excel.

El sistema genera automáticamente un dashboard ejecutivo que permite visualizar el desempeño de producción por día, semana y mes.

---

## Objetivos

- Monitorear cumplimiento del plan de producción.
- Analizar eficiencia operativa.
- Dar seguimiento al scrap.
- Controlar tiempo muerto.
- Medir productividad por operador.
- Generar reportes visuales para supervisión y toma de decisiones.

---

## Tecnologías utilizadas

- Python
- Pandas
- NumPy
- XlsxWriter
- OpenPyXL
- Microsoft Excel

---

## Indicadores (KPIs)

### Producción Total
Cantidad total producida durante el periodo.

### Cumplimiento del Plan (%)
Producción Real / Producción Plan.

### Eficiencia (%)
Rendimiento operativo de la línea.

### Scrap (%)
Porcentaje de material rechazado.

### Productividad por Operador
Producción Real / Número de Operadores.

### Tiempo Muerto (h)
Horas perdidas por paros o incidencias.

---

## Funcionalidades

- Análisis diario.
- Análisis semanal.
- Análisis mensual.
- Comparativo por turno.
- Comparativo por línea.
- Dashboard automático en Excel.
- Generación automática mediante Python.

---

## Estructura del Proyecto

```text
PRODUCCION/
│
├── Dashboard_Produccion_Diario_Semanal_Mensual.py
├── Dashboard_Produccion_Diario_Semanal_Mensual_V2.xlsx
├── Base_Produccion_Electronica.xlsx
└── Archivo_antiguo/
```

## Cómo ejecutar

Instalar dependencias:

```bash
py -m pip install pandas numpy xlsxwriter openpyxl
```

Ejecutar:

```bash
py Dashboard_Produccion_Diario_Semanal_Mensual.py
```

El sistema generará automáticamente:

```text
Dashboard_Produccion_Diario_Semanal_Mensual_V2.xlsx
```

---

## Autor

Ricardo Landa

Ingeniero Industrial enfocado en:

- Manufactura Electrónica
- Supervisión de Producción
- Mejora Continua
- Análisis de Datos
- Automatización con Python
