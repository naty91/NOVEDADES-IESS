
import io
import os
import re
import json
import zipfile
from datetime import date, datetime, timedelta
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st
import xlsxwriter
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
)

# ============================================================
# CONFIGURACIÓN GENERAL
# ============================================================

st.set_page_config(
    page_title="CENASE | Bitácora IESS",
    page_icon="🛡️",
    layout="wide",
)

TZ_ECUADOR = ZoneInfo("America/Guayaquil")
HOY = datetime.now(TZ_ECUADOR).date()

APP_VERSION = "1.0.0"

TIPOS_NOVEDAD = [
    "Aviso de entrada",
    "Aviso de salida – Renuncia",
    "Aviso de salida – Fin de contrato",
    "Aviso de salida – Terminación",
    "Aviso de salida – Abandono",
    "Modificación de sueldo",
    "Enfermedad",
    "Corrección de días",
    "Otra novedad",
]

INCIDENCIAS = [
    "Sin novedad",
    "Sistema IESS no habilitado",
    "Error/indisponibilidad IESS",
    "Reporte tardío de Operaciones",
    "Reporte tardío de RR. HH.",
    "Documentación incompleta",
    "Renuncia comunicada tardíamente",
    "Fecha no permitida por sistema",
    "Requiere regularización IESS",
    "Reporte tardío de Operaciones + Sistema IESS no habilitado",
    "Otra",
]

RESPONSABLES_ORIGEN = [
    "IESS",
    "OPERACIONES",
    "RR. HH.",
    "TRABAJADOR",
    "IESS + OPERACIONES",
    "IESS + RR. HH.",
    "OTRO",
]

ESTADOS = [
    "PENDIENTE",
    "EN TRÁMITE",
    "CERRADO",
    "REGULARIZACIÓN IESS",
]

EVIDENCIA_OPCIONES = ["NO", "SI"]

COLUMNAS = [
    "N.º",
    "Cédula",
    "Apellidos y nombres",
    "Cargo",
    "Centro de costo",
    "Tipo novedad",
    "Fecha efectiva",
    "Fecha reportada a RR.HH.",
    "Fecha intento IESS",
    "Situación / incidencia",
    "Responsable origen",
    "Evidencia",
    "Fecha registro IESS",
    "Fecha registrada IESS",
    "N.º trámite IESS",
    "Estado",
    "Observación",
    "DÍAS TRANSCURRIDOS",
    "ALERTA",
    "Archivo evidencia",
]

IESS_INCIDENCIAS = {
    "Sistema IESS no habilitado",
    "Error/indisponibilidad IESS",
    "Fecha no permitida por sistema",
    "Requiere regularización IESS",
    "Reporte tardío de Operaciones + Sistema IESS no habilitado",
}

# ============================================================
# ESTADO DE SESIÓN
# ============================================================

def init_state():
    if "registros" not in st.session_state:
        st.session_state.registros = []
    if "evidencias" not in st.session_state:
        # dict: id_registro -> {"name": str, "data": bytes}
        st.session_state.evidencias = {}
    if "config" not in st.session_state:
        st.session_state.config = {
            "empresa": "CENASE CÍA. LTDA.",
            "ruc": "0991317791001",
            "mes_control": HOY.strftime("%Y-%m"),
            "fecha_cierre_iess": None,
            "fecha_reapertura_iess": None,
            "responsable_elaboracion": "",
            "responsable_revision": "",
        }

init_state()

# ============================================================
# UTILIDADES
# ============================================================

def clean_text(value):
    if value is None:
        return ""
    if isinstance(value, float) and pd.isna(value):
        return ""
    return str(value).strip()

def normalize_cedula(value):
    txt = re.sub(r"\D", "", clean_text(value))
    return txt

def parse_date(value):
    if value in ("", None):
        return None
    if isinstance(value, date) and not isinstance(value, datetime):
        return value
    if isinstance(value, datetime):
        return value.date()
    try:
        ts = pd.to_datetime(value, errors="coerce")
        if pd.isna(ts):
            return None
        return ts.date()
    except Exception:
        return None

def fmt_date(value):
    d = parse_date(value)
    return d.strftime("%d/%m/%Y") if d else ""

def plazo_control(tipo):
    """Plazo de control usado por la app.
    - Entrada: 15 días.
    - Salida, enfermedad y modificación de sueldo: 3 días.
    - Otras: sin plazo automático.
    """
    tipo = clean_text(tipo)
    if tipo == "Aviso de entrada":
        return 15
    if tipo.startswith("Aviso de salida"):
        return 3
    if tipo in ("Enfermedad", "Modificación de sueldo"):
        return 3
    return None

def incidencia_iess(incidencia):
    return clean_text(incidencia) in IESS_INCIDENCIAS

def dias_transcurridos(fecha_efectiva, fecha_registro):
    fe = parse_date(fecha_efectiva)
    fr = parse_date(fecha_registro)
    if not fe:
        return None
    if fr:
        return (fr - fe).days
    return (HOY - fe).days

def calcular_alerta(reg):
    tipo = clean_text(reg.get("Tipo novedad"))
    plazo = plazo_control(tipo)
    fe = parse_date(reg.get("Fecha efectiva"))
    fr = parse_date(reg.get("Fecha registro IESS"))
    incidencia = clean_text(reg.get("Situación / incidencia"))
    estado = clean_text(reg.get("Estado")).upper()

    if not fe:
        return "⚪ SIN FECHA EFECTIVA"

    if plazo is None:
        if estado == "CERRADO":
            return "🔵 CERRADO - REVISIÓN MANUAL"
        return "🔵 REVISAR - SIN PLAZO AUTOMÁTICO"

    dias = dias_transcurridos(fe, fr)

    if fr:
        if dias <= plazo:
            return "🟢 DENTRO DEL PLAZO"
        if incidencia_iess(incidencia):
            return "🟡 FUERA DE PLAZO - INCIDENCIA IESS"
        return "🔴 FUERA DE PLAZO - REVISAR RESPONSABILIDAD"

    # Pendiente
    if dias <= plazo:
        return f"🟢 PENDIENTE - QUEDAN {plazo - dias} DÍA(S)"
    if incidencia_iess(incidencia):
        return "🟡 PENDIENTE FUERA DE PLAZO - INCIDENCIA IESS"
    return "🔴 PENDIENTE FUERA DE PLAZO - ACCIÓN INMEDIATA"

def sugerir_incidencia(fecha_efectiva, fecha_reportada, fecha_cierre, fecha_reapertura):
    fe = parse_date(fecha_efectiva)
    frep = parse_date(fecha_reportada)
    fc = parse_date(fecha_cierre)
    fra = parse_date(fecha_reapertura)

    if not fe:
        return None

    en_cierre = False
    if fc and fra:
        en_cierre = fc <= fe < fra

    reporte_tardio = bool(frep and frep > fe)

    if en_cierre and reporte_tardio:
        return "Reporte tardío de Operaciones + Sistema IESS no habilitado"
    if en_cierre:
        return "Sistema IESS no habilitado"
    if reporte_tardio:
        return "Reporte tardío de Operaciones"
    return "Sin novedad"

def next_numero():
    if not st.session_state.registros:
        return 1
    nums = []
    for r in st.session_state.registros:
        try:
            nums.append(int(r.get("N.º", 0)))
        except Exception:
            pass
    return max(nums or [0]) + 1

def make_id(reg):
    ced = normalize_cedula(reg.get("Cédula"))
    fe = fmt_date(reg.get("Fecha efectiva")).replace("/", "")
    n = reg.get("N.º", "")
    return f"{n}_{ced}_{fe}"

def dataframe_actual():
    rows = []
    for r in st.session_state.registros:
        rr = dict(r)
        rr["DÍAS TRANSCURRIDOS"] = dias_transcurridos(
            rr.get("Fecha efectiva"), rr.get("Fecha registro IESS")
        )
        rr["ALERTA"] = calcular_alerta(rr)
        evidence = st.session_state.evidencias.get(make_id(rr))
        rr["Archivo evidencia"] = evidence["name"] if evidence else clean_text(rr.get("Archivo evidencia"))
        rows.append(rr)
    df = pd.DataFrame(rows)
    for col in COLUMNAS:
        if col not in df.columns:
            df[col] = ""
    return df[COLUMNAS]

def record_from_row(row):
    reg = {}
    for col in COLUMNAS:
        if col in ("DÍAS TRANSCURRIDOS", "ALERTA", "Archivo evidencia"):
            continue
        val = row.get(col, "")
        if col.startswith("Fecha"):
            val = parse_date(val)
        else:
            val = clean_text(val)
        reg[col] = val
    try:
        reg["N.º"] = int(float(row.get("N.º")))
    except Exception:
        reg["N.º"] = next_numero()
    reg["Cédula"] = normalize_cedula(reg.get("Cédula"))
    return reg

# ============================================================
# IMPORTAR EXCEL EXISTENTE
# ============================================================

def importar_excel(uploaded):
    data = uploaded.getvalue()
    xls = pd.ExcelFile(io.BytesIO(data))
    sheet = "BITACORA" if "BITACORA" in xls.sheet_names else xls.sheet_names[0]
    df = pd.read_excel(io.BytesIO(data), sheet_name=sheet)

    # Normalización de encabezados equivalentes
    rename = {}
    for c in df.columns:
        key = clean_text(c).upper()
        if key == "BITÁCORA":
            continue
    df = df.dropna(how="all")
    nuevos = []
    for _, row in df.iterrows():
        if not clean_text(row.get("Cédula", "")):
            continue
        nuevos.append(record_from_row(row))
    return nuevos

# ============================================================
# EXCEL
# ============================================================

def excel_bytes(df, config):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter", datetime_format="dd/mm/yyyy") as writer:
        export = df.copy()
        for col in ["Fecha efectiva", "Fecha reportada a RR.HH.", "Fecha intento IESS",
                    "Fecha registro IESS", "Fecha registrada IESS"]:
            if col in export.columns:
                export[col] = pd.to_datetime(export[col], errors="coerce")
        export.to_excel(writer, sheet_name="BITACORA", index=False, startrow=5)

        workbook = writer.book
        ws = writer.sheets["BITACORA"]
        ws.freeze_panes(6, 2)
        ws.autofilter(5, 0, 5 + len(export), len(export.columns)-1)

        title_fmt = workbook.add_format({
            "bold": True, "font_size": 16, "align": "center",
            "valign": "vcenter", "bg_color": "#1F4E78", "font_color": "white"
        })
        sub_fmt = workbook.add_format({"bold": True, "font_size": 10})
        header_fmt = workbook.add_format({
            "bold": True, "bg_color": "#D9EAF7", "border": 1,
            "align": "center", "valign": "vcenter", "text_wrap": True
        })
        wrap_fmt = workbook.add_format({"text_wrap": True, "valign": "top", "border": 1})
        center_fmt = workbook.add_format({"align": "center", "valign": "top", "border": 1})
        date_fmt = workbook.add_format({"num_format": "dd/mm/yyyy", "align": "center", "border": 1})
        green_fmt = workbook.add_format({"bg_color": "#E2F0D9"})
        yellow_fmt = workbook.add_format({"bg_color": "#FFF2CC"})
        red_fmt = workbook.add_format({"bg_color": "#F4CCCC"})

        ws.merge_range(0, 0, 0, len(export.columns)-1, "CENASE CÍA. LTDA. - BITÁCORA DE NOVEDADES E INCIDENCIAS IESS", title_fmt)
        ws.write(1, 0, "RUC:", sub_fmt); ws.write(1, 1, config.get("ruc",""))
        ws.write(2, 0, "Mes de control:", sub_fmt); ws.write(2, 1, config.get("mes_control",""))
        ws.write(2, 3, "Cierre IESS:", sub_fmt); ws.write(2, 4, fmt_date(config.get("fecha_cierre_iess")))
        ws.write(2, 6, "Reapertura IESS:", sub_fmt); ws.write(2, 7, fmt_date(config.get("fecha_reapertura_iess")))
        ws.write(3, 0, "Elaborado por:", sub_fmt); ws.write(3, 1, config.get("responsable_elaboracion",""))
        ws.write(3, 3, "Revisado por:", sub_fmt); ws.write(3, 4, config.get("responsable_revision",""))

        for c, col in enumerate(export.columns):
            ws.write(5, c, col, header_fmt)

        widths = {
            0: 7, 1: 14, 2: 34, 3: 20, 4: 24, 5: 32,
            6: 14, 7: 18, 8: 15, 9: 38, 10: 23, 11: 12,
            12: 15, 13: 17, 14: 18, 15: 18, 16: 45, 17: 18, 18: 42, 19: 26
        }
        for c in range(len(export.columns)):
            ws.set_column(c, c, widths.get(c, 18))

        date_cols = {export.columns.get_loc(c) for c in export.columns if c.startswith("Fecha")}
        for r in range(len(export)):
            for c in range(len(export.columns)):
                val = export.iloc[r, c]
                rr = 6 + r
                if pd.isna(val):
                    ws.write_blank(rr, c, None, center_fmt if c in date_cols else wrap_fmt)
                elif c in date_cols and isinstance(val, (pd.Timestamp, datetime)):
                    ws.write_datetime(rr, c, val.to_pydatetime(), date_fmt)
                elif c in (0, 1, 6, 7, 8, 11, 12, 13, 14, 15, 17):
                    ws.write(rr, c, val, center_fmt)
                else:
                    ws.write(rr, c, val, wrap_fmt)

        if len(export):
            alert_col = export.columns.get_loc("ALERTA")
            first = 6
            last = 5 + len(export)
            ws.conditional_format(first, alert_col, last, alert_col, {
                "type": "text", "criteria": "containing", "value": "🟢", "format": green_fmt
            })
            ws.conditional_format(first, alert_col, last, alert_col, {
                "type": "text", "criteria": "containing", "value": "🟡", "format": yellow_fmt
            })
            ws.conditional_format(first, alert_col, last, alert_col, {
                "type": "text", "criteria": "containing", "value": "🔴", "format": red_fmt
            })

        # Hoja de resumen
        rs = workbook.add_worksheet("RESUMEN")
        rs.set_column("A:A", 38)
        rs.set_column("B:B", 18)
        rs.write("A1", "RESUMEN DE CONTROL IESS", title_fmt)
        metricas = [
            ("Total novedades", len(export)),
            ("Dentro del plazo", int(export["ALERTA"].astype(str).str.contains("🟢").sum()) if len(export) else 0),
            ("Incidencia IESS", int(export["ALERTA"].astype(str).str.contains("🟡").sum()) if len(export) else 0),
            ("Revisar responsabilidad", int(export["ALERTA"].astype(str).str.contains("🔴").sum()) if len(export) else 0),
            ("Pendientes", int(export["Estado"].astype(str).str.upper().eq("PENDIENTE").sum()) if len(export) else 0),
            ("En regularización", int(export["Estado"].astype(str).str.upper().eq("REGULARIZACIÓN IESS").sum()) if len(export) else 0),
        ]
        for i, (k,v) in enumerate(metricas, start=3):
            rs.write(i-1, 0, k, sub_fmt)
            rs.write(i-1, 1, v)

    return output.getvalue()

# ============================================================
# PDF
# ============================================================

def pdf_bytes(df, config, titulo="Reporte de Novedades e Incidencias IESS"):
    output = io.BytesIO()
    doc = SimpleDocTemplate(
        output, pagesize=landscape(A4),
        rightMargin=8*mm, leftMargin=8*mm, topMargin=10*mm, bottomMargin=10*mm
    )
    styles = getSampleStyleSheet()
    title = ParagraphStyle(
        "TitleCENASE", parent=styles["Title"], fontSize=14, leading=16,
        alignment=TA_CENTER, textColor=colors.HexColor("#1F4E78")
    )
    small = ParagraphStyle("small", parent=styles["Normal"], fontSize=6.6, leading=8)
    small_center = ParagraphStyle("smallcenter", parent=small, alignment=TA_CENTER)
    story = [
        Paragraph(f"<b>{config.get('empresa','CENASE CÍA. LTDA.')}</b>", title),
        Paragraph(titulo, title),
        Spacer(1, 4*mm),
    ]

    meta = (
        f"<b>RUC:</b> {config.get('ruc','')} &nbsp;&nbsp;&nbsp; "
        f"<b>Mes:</b> {config.get('mes_control','')} &nbsp;&nbsp;&nbsp; "
        f"<b>Cierre IESS:</b> {fmt_date(config.get('fecha_cierre_iess')) or 'No definido'} &nbsp;&nbsp;&nbsp; "
        f"<b>Reapertura:</b> {fmt_date(config.get('fecha_reapertura_iess')) or 'No definida'}"
    )
    story.append(Paragraph(meta, styles["Normal"]))
    story.append(Spacer(1, 4*mm))

    cols = [
        "N.º", "Cédula", "Apellidos y nombres", "Centro de costo",
        "Tipo novedad", "Fecha efectiva", "Fecha reportada a RR.HH.",
        "Situación / incidencia", "Fecha registro IESS", "Estado", "ALERTA"
    ]
    data = [[Paragraph(f"<b>{c}</b>", small_center) for c in cols]]
    for _, row in df.iterrows():
        vals = []
        for c in cols:
            v = row.get(c, "")
            if c.startswith("Fecha"):
                v = fmt_date(v)
            vals.append(Paragraph(clean_text(v), small_center if c in ["N.º","Cédula","Fecha efectiva","Fecha reportada a RR.HH.","Fecha registro IESS","Estado"] else small))
        data.append(vals)

    widths = [10*mm, 21*mm, 40*mm, 29*mm, 42*mm, 20*mm, 24*mm, 47*mm, 22*mm, 22*mm, 50*mm]
    table = Table(data, colWidths=widths, repeatRows=1)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#D9EAF7")),
        ("TEXTCOLOR", (0,0), (-1,0), colors.black),
        ("GRID", (0,0), (-1,-1), 0.35, colors.HexColor("#A6A6A6")),
        ("VALIGN", (0,0), (-1,-1), "TOP"),
        ("LEFTPADDING", (0,0), (-1,-1), 2),
        ("RIGHTPADDING", (0,0), (-1,-1), 2),
        ("TOPPADDING", (0,0), (-1,-1), 2),
        ("BOTTOMPADDING", (0,0), (-1,-1), 2),
    ]))
    story.append(table)
    story.append(Spacer(1, 4*mm))

    counts = {
        "Total": len(df),
        "Dentro del plazo": int(df["ALERTA"].astype(str).str.contains("🟢").sum()) if len(df) else 0,
        "Incidencia IESS": int(df["ALERTA"].astype(str).str.contains("🟡").sum()) if len(df) else 0,
        "Revisar responsabilidad": int(df["ALERTA"].astype(str).str.contains("🔴").sum()) if len(df) else 0,
    }
    story.append(Paragraph(
        " | ".join(f"<b>{k}:</b> {v}" for k,v in counts.items()),
        styles["Normal"]
    ))
    story.append(Spacer(1, 3*mm))
    story.append(Paragraph(
        "<b>Nota de control:</b> la app conserva separadas la fecha efectiva de la novedad y la fecha en que pudo registrarse en IESS. "
        "No se debe modificar la fecha laboral real para hacerla coincidir con la disponibilidad del portal.",
        styles["Normal"]
    ))

    doc.build(story)
    return output.getvalue()

def ficha_pdf_bytes(reg, config):
    df = pd.DataFrame([{**reg,
                        "DÍAS TRANSCURRIDOS": dias_transcurridos(reg.get("Fecha efectiva"), reg.get("Fecha registro IESS")),
                        "ALERTA": calcular_alerta(reg)}])
    # PDF detallado individual
    output = io.BytesIO()
    doc = SimpleDocTemplate(output, pagesize=A4, rightMargin=18*mm, leftMargin=18*mm, topMargin=15*mm, bottomMargin=15*mm)
    styles = getSampleStyleSheet()
    title = ParagraphStyle("title2", parent=styles["Title"], fontSize=14, textColor=colors.HexColor("#1F4E78"))
    story = [
        Paragraph("<b>CENASE CÍA. LTDA.</b>", title),
        Paragraph("Ficha individual de novedad / incidencia IESS", title),
        Spacer(1, 5*mm),
    ]
    campos = [c for c in COLUMNAS if c != "Archivo evidencia"]
    data = []
    for c in campos:
        v = reg.get(c, "")
        if c == "DÍAS TRANSCURRIDOS":
            v = dias_transcurridos(reg.get("Fecha efectiva"), reg.get("Fecha registro IESS"))
        elif c == "ALERTA":
            v = calcular_alerta(reg)
        elif c.startswith("Fecha"):
            v = fmt_date(v)
        data.append([Paragraph(f"<b>{c}</b>", styles["BodyText"]), Paragraph(clean_text(v), styles["BodyText"])])
    t = Table(data, colWidths=[60*mm, 110*mm])
    t.setStyle(TableStyle([
        ("GRID",(0,0),(-1,-1),0.4,colors.HexColor("#A6A6A6")),
        ("BACKGROUND",(0,0),(0,-1),colors.HexColor("#EAF2F8")),
        ("VALIGN",(0,0),(-1,-1),"TOP"),
        ("LEFTPADDING",(0,0),(-1,-1),4),
        ("RIGHTPADDING",(0,0),(-1,-1),4),
        ("TOPPADDING",(0,0),(-1,-1),4),
        ("BOTTOMPADDING",(0,0),(-1,-1),4),
    ]))
    story.append(t)
    story.append(Spacer(1, 6*mm))
    story.append(Paragraph(
        "<b>Control interno:</b> Mantener adjunto el soporte laboral, evidencia de indisponibilidad del sistema cuando corresponda, "
        "comprobante de registro y número de trámite IESS si existe.",
        styles["BodyText"]
    ))
    doc.build(story)
    return output.getvalue()

# ============================================================
# RESPALDO ZIP
# ============================================================

def backup_zip_bytes():
    buf = io.BytesIO()
    config = dict(st.session_state.config)
    for k in ["fecha_cierre_iess", "fecha_reapertura_iess"]:
        if isinstance(config.get(k), date):
            config[k] = config[k].isoformat()

    registros = []
    for r in st.session_state.registros:
        rr = dict(r)
        for k, v in rr.items():
            if isinstance(v, (date, datetime)):
                rr[k] = v.isoformat()
        registros.append(rr)

    manifest = {
        "app": "CENASE Bitácora IESS",
        "version": APP_VERSION,
        "fecha_respaldo": datetime.now(TZ_ECUADOR).isoformat(),
        "config": config,
        "registros": registros,
    }

    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("respaldo.json", json.dumps(manifest, ensure_ascii=False, indent=2))
        for rec_id, item in st.session_state.evidencias.items():
            safe = re.sub(r"[^A-Za-z0-9._-]", "_", item["name"])
            z.writestr(f"evidencias/{rec_id}__{safe}", item["data"])
    return buf.getvalue()

def restore_backup(uploaded):
    data = uploaded.getvalue()
    with zipfile.ZipFile(io.BytesIO(data), "r") as z:
        manifest = json.loads(z.read("respaldo.json").decode("utf-8"))
        cfg = manifest.get("config", {})
        for k in ["fecha_cierre_iess", "fecha_reapertura_iess"]:
            cfg[k] = parse_date(cfg.get(k))
        regs = []
        for r in manifest.get("registros", []):
            rr = dict(r)
            for k in list(rr):
                if k.startswith("Fecha"):
                    rr[k] = parse_date(rr[k])
            regs.append(rr)

        evid = {}
        for name in z.namelist():
            if not name.startswith("evidencias/") or name.endswith("/"):
                continue
            base = name.split("/",1)[1]
            if "__" not in base:
                continue
            rec_id, filename = base.split("__",1)
            evid[rec_id] = {"name": filename, "data": z.read(name)}

    st.session_state.config.update(cfg)
    st.session_state.registros = regs
    st.session_state.evidencias = evid

# ============================================================
# ESTILOS
# ============================================================

st.markdown("""
<style>
.block-container {padding-top: 1.1rem; padding-bottom: 2rem;}
div[data-testid="stMetric"] {
    background: #f7f9fb; border: 1px solid #dde5ec; padding: 10px; border-radius: 10px;
}
.small-note {font-size: 0.86rem; color: #555;}
</style>
""", unsafe_allow_html=True)

# ============================================================
# ENCABEZADO
# ============================================================

st.title("🛡️ CENASE | Control de Novedades e Incidencias IESS")
st.caption(
    "Bitácora operativa para RR. HH. + conciliación de Contabilidad. "
    "Distingue fecha laboral real, fecha de reporte interno, intento y registro efectivo en IESS."
)

# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:
    st.header("⚙️ Configuración del período")
    cfg = st.session_state.config

    cfg["mes_control"] = st.text_input("Mes de control (AAAA-MM)", value=cfg.get("mes_control",""))
    cierre_default = parse_date(cfg.get("fecha_cierre_iess"))
    reap_default = parse_date(cfg.get("fecha_reapertura_iess"))

    usar_cierre = st.checkbox("Registrar período de cierre IESS", value=bool(cierre_default or reap_default))
    if usar_cierre:
        cfg["fecha_cierre_iess"] = st.date_input(
            "Fecha de cierre / indisponibilidad",
            value=cierre_default or HOY,
            format="DD/MM/YYYY"
        )
        cfg["fecha_reapertura_iess"] = st.date_input(
            "Fecha de reapertura",
            value=reap_default or HOY,
            format="DD/MM/YYYY"
        )
    else:
        cfg["fecha_cierre_iess"] = None
        cfg["fecha_reapertura_iess"] = None

    cfg["responsable_elaboracion"] = st.text_input(
        "Responsable elaboración", value=cfg.get("responsable_elaboracion","")
    )
    cfg["responsable_revision"] = st.text_input(
        "Responsable revisión", value=cfg.get("responsable_revision","")
    )

    st.divider()
    st.subheader("📥 Cargar bitácora Excel")
    excel_up = st.file_uploader("Excel existente", type=["xlsx"], key="excel_import")
    if excel_up and st.button("Importar registros del Excel", use_container_width=True):
        try:
            nuevos = importar_excel(excel_up)
            st.session_state.registros = nuevos
            st.success(f"Se importaron {len(nuevos)} registros.")
            st.rerun()
        except Exception as e:
            st.error(f"No se pudo importar el Excel: {e}")

    st.divider()
    st.subheader("💾 Respaldo")
    respaldo_up = st.file_uploader("Subir respaldo (.zip)", type=["zip"], key="backup_import")
    if respaldo_up and st.button("Restaurar respaldo", use_container_width=True):
        try:
            restore_backup(respaldo_up)
            st.success("Respaldo restaurado.")
            st.rerun()
        except Exception as e:
            st.error(f"Respaldo inválido o dañado: {e}")

    backup = backup_zip_bytes()
    st.download_button(
        "⬇️ Bajar respaldo completo",
        data=backup,
        file_name=f"CENASE_IESS_respaldo_{HOY.isoformat()}.zip",
        mime="application/zip",
        use_container_width=True
    )

    st.divider()
    st.caption(f"Versión {APP_VERSION}")

# ============================================================
# DASHBOARD
# ============================================================

df = dataframe_actual()

m1, m2, m3, m4, m5 = st.columns(5)
m1.metric("Total novedades", len(df))
m2.metric("🟢 Dentro de plazo", int(df["ALERTA"].astype(str).str.contains("🟢").sum()) if len(df) else 0)
m3.metric("🟡 Incidencia IESS", int(df["ALERTA"].astype(str).str.contains("🟡").sum()) if len(df) else 0)
m4.metric("🔴 Revisar", int(df["ALERTA"].astype(str).str.contains("🔴").sum()) if len(df) else 0)
m5.metric("Pendientes", int(df["Estado"].astype(str).str.upper().eq("PENDIENTE").sum()) if len(df) else 0)

tabs = st.tabs(["➕ Registrar novedad", "📋 Bitácora y edición", "📊 Reportes", "📚 Criterios de control"])

# ============================================================
# TAB 1: REGISTRAR
# ============================================================

with tabs[0]:
    st.subheader("Registrar nueva novedad")
    st.info(
        "Regla CENASE: la fecha efectiva siempre debe reflejar la realidad laboral. "
        "No se modifica para hacerla coincidir con la fecha en que el IESS habilita el sistema."
    )

    with st.form("form_nueva", clear_on_submit=True):
        c1, c2, c3 = st.columns(3)
        cedula = c1.text_input("Cédula *", max_chars=13)
        nombres = c2.text_input("Apellidos y nombres *")
        cargo = c3.text_input("Cargo", value="GUARDIA")

        c4, c5, c6 = st.columns(3)
        centro = c4.text_input("Centro de costo")
        tipo = c5.selectbox("Tipo de novedad *", TIPOS_NOVEDAD)
        fecha_ef = c6.date_input("Fecha efectiva *", value=HOY, format="DD/MM/YYYY")

        c7, c8, c9 = st.columns(3)
        fecha_rep = c7.date_input("Fecha reportada a RR. HH.", value=HOY, format="DD/MM/YYYY")
        fecha_int = c8.date_input("Fecha intento IESS", value=HOY, format="DD/MM/YYYY")
        sugerencia = sugerir_incidencia(
            fecha_ef, fecha_rep,
            st.session_state.config.get("fecha_cierre_iess"),
            st.session_state.config.get("fecha_reapertura_iess")
        )
        idx_inc = INCIDENCIAS.index(sugerencia) if sugerencia in INCIDENCIAS else 0
        incidencia = c9.selectbox("Situación / incidencia", INCIDENCIAS, index=idx_inc)

        c10, c11, c12 = st.columns(3)
        resp_origen = c10.selectbox("Responsable origen", RESPONSABLES_ORIGEN)
        evidencia = c11.selectbox("¿Existe evidencia?", EVIDENCIA_OPCIONES, index=0)
        estado = c12.selectbox("Estado", ESTADOS, index=0)

        c13, c14, c15 = st.columns(3)
        registrar_ahora = c13.checkbox("Ya fue registrado en IESS")
        fecha_reg = None
        fecha_reg_iess = None
        if registrar_ahora:
            fecha_reg = c14.date_input("Fecha de registro IESS", value=HOY, format="DD/MM/YYYY")
            fecha_reg_iess = c15.date_input("Fecha registrada en IESS", value=fecha_ef, format="DD/MM/YYYY")

        c16, c17 = st.columns([1,2])
        tramite = c16.text_input("N.º trámite IESS")
        obs = c17.text_area("Observación", height=90)

        evidencia_file = st.file_uploader(
            "Adjuntar evidencia (opcional: PDF/JPG/PNG)",
            type=["pdf","jpg","jpeg","png"],
            key="evidencia_nueva"
        )

        enviar = st.form_submit_button("💾 Guardar novedad", use_container_width=True)

    if enviar:
        if not normalize_cedula(cedula) or not clean_text(nombres):
            st.error("Cédula y nombres son obligatorios.")
        else:
            reg = {
                "N.º": next_numero(),
                "Cédula": normalize_cedula(cedula),
                "Apellidos y nombres": nombres.strip().upper(),
                "Cargo": cargo.strip().upper(),
                "Centro de costo": centro.strip().upper(),
                "Tipo novedad": tipo,
                "Fecha efectiva": fecha_ef,
                "Fecha reportada a RR.HH.": fecha_rep,
                "Fecha intento IESS": fecha_int,
                "Situación / incidencia": incidencia,
                "Responsable origen": resp_origen,
                "Evidencia": evidencia,
                "Fecha registro IESS": fecha_reg,
                "Fecha registrada IESS": fecha_reg_iess,
                "N.º trámite IESS": tramite.strip().upper(),
                "Estado": "CERRADO" if registrar_ahora and estado == "PENDIENTE" else estado,
                "Observación": obs.strip(),
            }
            st.session_state.registros.append(reg)
            if evidencia_file:
                st.session_state.evidencias[make_id(reg)] = {
                    "name": evidencia_file.name,
                    "data": evidencia_file.getvalue(),
                }
                reg["Evidencia"] = "SI"
            st.success(f"Novedad N.º {reg['N.º']} registrada. Alerta: {calcular_alerta(reg)}")
            st.rerun()

# ============================================================
# TAB 2: BITÁCORA Y EDICIÓN
# ============================================================

with tabs[1]:
    st.subheader("Bitácora")

    if df.empty:
        st.warning("Todavía no existen registros.")
    else:
        f1, f2, f3, f4 = st.columns(4)
        filtro_estado = f1.multiselect("Estado", sorted(df["Estado"].dropna().astype(str).unique()))
        filtro_tipo = f2.multiselect("Tipo de novedad", sorted(df["Tipo novedad"].dropna().astype(str).unique()))
        filtro_cc = f3.multiselect("Centro de costo", sorted([x for x in df["Centro de costo"].dropna().astype(str).unique() if x]))
        filtro_alerta = f4.selectbox("Alerta", ["Todas","🟢","🟡","🔴"])

        vista = df.copy()
        if filtro_estado:
            vista = vista[vista["Estado"].isin(filtro_estado)]
        if filtro_tipo:
            vista = vista[vista["Tipo novedad"].isin(filtro_tipo)]
        if filtro_cc:
            vista = vista[vista["Centro de costo"].isin(filtro_cc)]
        if filtro_alerta != "Todas":
            vista = vista[vista["ALERTA"].astype(str).str.contains(filtro_alerta, regex=False)]

        st.dataframe(
            vista.drop(columns=["Archivo evidencia"], errors="ignore"),
            use_container_width=True,
            hide_index=True,
            height=430
        )

        st.divider()
        st.subheader("Editar / cerrar / eliminar registro")
        opciones = {
            f"{int(r['N.º'])} | {r['Cédula']} | {r['Apellidos y nombres']} | {fmt_date(r['Fecha efectiva'])}": i
            for i, r in enumerate(st.session_state.registros)
        }
        sel = st.selectbox("Seleccione un registro", list(opciones.keys()))
        idx = opciones[sel]
        reg = st.session_state.registros[idx]

        with st.form("form_editar"):
            e1, e2, e3 = st.columns(3)
            e_ced = e1.text_input("Cédula", value=clean_text(reg.get("Cédula")))
            e_nom = e2.text_input("Apellidos y nombres", value=clean_text(reg.get("Apellidos y nombres")))
            e_cargo = e3.text_input("Cargo", value=clean_text(reg.get("Cargo")))

            e4, e5, e6 = st.columns(3)
            e_cc = e4.text_input("Centro de costo", value=clean_text(reg.get("Centro de costo")))
            e_tipo = e5.selectbox("Tipo novedad", TIPOS_NOVEDAD, index=TIPOS_NOVEDAD.index(reg.get("Tipo novedad")) if reg.get("Tipo novedad") in TIPOS_NOVEDAD else 0)
            e_fe = e6.date_input("Fecha efectiva", value=parse_date(reg.get("Fecha efectiva")) or HOY, format="DD/MM/YYYY")

            e7, e8, e9 = st.columns(3)
            e_fr = e7.date_input("Fecha reportada a RR.HH.", value=parse_date(reg.get("Fecha reportada a RR.HH.")) or HOY, format="DD/MM/YYYY")
            e_fi = e8.date_input("Fecha intento IESS", value=parse_date(reg.get("Fecha intento IESS")) or HOY, format="DD/MM/YYYY")
            e_inc = e9.selectbox("Situación / incidencia", INCIDENCIAS, index=INCIDENCIAS.index(reg.get("Situación / incidencia")) if reg.get("Situación / incidencia") in INCIDENCIAS else 0)

            e10, e11, e12 = st.columns(3)
            e_resp = e10.selectbox("Responsable origen", RESPONSABLES_ORIGEN, index=RESPONSABLES_ORIGEN.index(reg.get("Responsable origen")) if reg.get("Responsable origen") in RESPONSABLES_ORIGEN else 0)
            e_evid = e11.selectbox("Evidencia", EVIDENCIA_OPCIONES, index=1 if clean_text(reg.get("Evidencia")).upper()=="SI" else 0)
            e_estado = e12.selectbox("Estado", ESTADOS, index=ESTADOS.index(reg.get("Estado")) if reg.get("Estado") in ESTADOS else 0)

            e13, e14, e15 = st.columns(3)
            has_freg = e13.checkbox("Tiene registro IESS", value=bool(parse_date(reg.get("Fecha registro IESS"))))
            e_freg = e14.date_input("Fecha registro IESS", value=parse_date(reg.get("Fecha registro IESS")) or HOY, format="DD/MM/YYYY", disabled=not has_freg)
            e_fregi = e15.date_input("Fecha registrada IESS", value=parse_date(reg.get("Fecha registrada IESS")) or e_fe, format="DD/MM/YYYY", disabled=not has_freg)

            e16, e17 = st.columns([1,2])
            e_tram = e16.text_input("N.º trámite", value=clean_text(reg.get("N.º trámite IESS")))
            e_obs = e17.text_area("Observación", value=clean_text(reg.get("Observación")), height=90)

            new_ev = st.file_uploader("Reemplazar / adjuntar evidencia", type=["pdf","jpg","jpeg","png"], key=f"ev_edit_{idx}")

            guardar = st.form_submit_button("Guardar cambios", use_container_width=True)

        if guardar:
            old_id = make_id(reg)
            actual = {
                "N.º": reg["N.º"],
                "Cédula": normalize_cedula(e_ced),
                "Apellidos y nombres": e_nom.strip().upper(),
                "Cargo": e_cargo.strip().upper(),
                "Centro de costo": e_cc.strip().upper(),
                "Tipo novedad": e_tipo,
                "Fecha efectiva": e_fe,
                "Fecha reportada a RR.HH.": e_fr,
                "Fecha intento IESS": e_fi,
                "Situación / incidencia": e_inc,
                "Responsable origen": e_resp,
                "Evidencia": e_evid,
                "Fecha registro IESS": e_freg if has_freg else None,
                "Fecha registrada IESS": e_fregi if has_freg else None,
                "N.º trámite IESS": e_tram.strip().upper(),
                "Estado": e_estado,
                "Observación": e_obs.strip(),
            }
            new_id = make_id(actual)
            if old_id != new_id and old_id in st.session_state.evidencias:
                st.session_state.evidencias[new_id] = st.session_state.evidencias.pop(old_id)
            if new_ev:
                st.session_state.evidencias[new_id] = {"name": new_ev.name, "data": new_ev.getvalue()}
                actual["Evidencia"] = "SI"
            st.session_state.registros[idx] = actual
            st.success(f"Actualizado. {calcular_alerta(actual)}")
            st.rerun()

        b1, b2 = st.columns(2)
        ficha = ficha_pdf_bytes(reg, st.session_state.config)
        b1.download_button(
            "📄 PDF individual",
            ficha,
            file_name=f"Novedad_IESS_{reg['N.º']}_{reg['Cédula']}.pdf",
            mime="application/pdf",
            use_container_width=True
        )
        if b2.button("🗑️ Eliminar registro", type="secondary", use_container_width=True):
            rid = make_id(reg)
            st.session_state.evidencias.pop(rid, None)
            st.session_state.registros.pop(idx)
            st.success("Registro eliminado.")
            st.rerun()

# ============================================================
# TAB 3: REPORTES
# ============================================================

with tabs[2]:
    st.subheader("Reportes descargables")

    if df.empty:
        st.warning("No hay información para exportar.")
    else:
        r1, r2, r3 = st.columns(3)
        rep_estado = r1.multiselect("Filtrar por estado", sorted(df["Estado"].astype(str).unique()), key="rep_estado")
        rep_cc = r2.multiselect("Filtrar por centro de costo", sorted([x for x in df["Centro de costo"].astype(str).unique() if x]), key="rep_cc")
        rep_alerta = r3.selectbox("Filtrar alerta", ["Todas","🟢","🟡","🔴"], key="rep_alerta")

        reporte = df.copy()
        if rep_estado:
            reporte = reporte[reporte["Estado"].isin(rep_estado)]
        if rep_cc:
            reporte = reporte[reporte["Centro de costo"].isin(rep_cc)]
        if rep_alerta != "Todas":
            reporte = reporte[reporte["ALERTA"].astype(str).str.contains(rep_alerta, regex=False)]

        st.caption(f"Registros incluidos en el reporte: {len(reporte)}")
        st.dataframe(reporte.drop(columns=["Archivo evidencia"], errors="ignore"), use_container_width=True, hide_index=True)

        xlsx = excel_bytes(reporte.drop(columns=["Archivo evidencia"], errors="ignore"), st.session_state.config)
        pdf = pdf_bytes(reporte.drop(columns=["Archivo evidencia"], errors="ignore"), st.session_state.config)

        d1, d2 = st.columns(2)
        d1.download_button(
            "📊 Descargar Excel",
            xlsx,
            file_name=f"CENASE_Bitacora_IESS_{st.session_state.config.get('mes_control','')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        d2.download_button(
            "📄 Descargar PDF",
            pdf,
            file_name=f"CENASE_Bitacora_IESS_{st.session_state.config.get('mes_control','')}.pdf",
            mime="application/pdf",
            use_container_width=True
        )

        st.markdown("#### Resumen por causa")
        causa = reporte.groupby("Situación / incidencia", dropna=False).size().reset_index(name="Cantidad").sort_values("Cantidad", ascending=False)
        st.dataframe(causa, use_container_width=True, hide_index=True)

        st.markdown("#### Resumen por responsable de origen")
        resp = reporte.groupby("Responsable origen", dropna=False).size().reset_index(name="Cantidad").sort_values("Cantidad", ascending=False)
        st.dataframe(resp, use_container_width=True, hide_index=True)

# ============================================================
# TAB 4: CRITERIOS
# ============================================================

with tabs[3]:
    st.subheader("Criterios operativos incorporados en la app")
    st.markdown("""
**1. Fecha laboral real.** La fecha efectiva de ingreso o salida no se cambia por el hecho de que Historia Laboral esté cerrado.

**2. Aviso de entrada.** La app usa 15 días como plazo de control.

**3. Avisos de salida, enfermedad y modificación de sueldo.** La app usa 3 días como plazo de control.

**4. Cierre o indisponibilidad IESS.** Si el plazo se supera y la causa registrada corresponde al IESS, la alerta se muestra en amarillo para separar la incidencia del sistema de una posible falla interna.

**5. Reporte interno tardío.** Si Operaciones o RR. HH. comunica/procesa tarde y no existe una incidencia IESS registrada, la alerta fuera de plazo es roja.

**6. Caso mixto.** La app permite registrar “Reporte tardío de Operaciones + Sistema IESS no habilitado” para no ocultar ninguna de las dos causas.

**7. Reapertura.** Las novedades pendientes deben gestionarse como prioridad cuando el sistema vuelva a estar disponible.

**8. Fecha no permitida por el portal.** No se inventa una nueva fecha; el caso pasa a regularización y debe conservarse el número de trámite y los soportes.

**9. Contabilidad.** La bitácora sirve como soporte para conciliar Rol vs IESS y explicar diferencias de días o valores. La responsabilidad de documentar la novedad laboral debe permanecer en RR. HH./Operaciones según corresponda.

**10. Respaldo.** Streamlit Community Cloud no debe usarse como único repositorio permanente. La app permite bajar un ZIP con registros + evidencias y restaurarlo después.
""")
    st.warning(
        "Los plazos se utilizan como reglas de control conforme a la información oficial revisada. "
        "La app no presume que el cierre del sistema suspenda automáticamente los plazos legales. "
        "Cuando CENASE reciba un pronunciamiento formal del IESS, estas reglas pueden actualizarse."
    )

    st.markdown("#### Semáforo")
    st.markdown("""
- 🟢 **Dentro del plazo / pendiente aún dentro del plazo**
- 🟡 **Fuera de plazo o pendiente, pero con incidencia IESS documentada**
- 🔴 **Fuera de plazo sin causa IESS registrada: revisar responsabilidad interna**
- 🔵 **Novedad sin plazo automático configurado**
""")
