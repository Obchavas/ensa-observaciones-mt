# -*- coding: utf-8 -*-
"""
ENSA Corporate Theme - Utilidades
CORRECCIÓN CRÍTICA: Zona UTM 17S (no 18M) para Lambayeque
"""
import io
import utm
import pandas as pd
import folium
from datetime import datetime
from fpdf import FPDF
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from pathlib import Path
import base64

# ============================================================
# PALETA CORPORATIVA
# ============================================================
BRAND = {
    "primary_dark_blue": "#003B70",
    "secondary_blue":    "#0072CE",
    "brand_green":       "#38B449",
    "text_main":         "#003B70",
    "text_secondary":    "#2C4D6F",
    "background_light":  "#FFFFFF",
}

FONT_TITLE = "Montserrat"
FONT_BODY  = "Open Sans"

def hex_rgb(hexcolor: str):
    h = hexcolor.lstrip("#")
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))

# ============================================================
# ⚠️ CORRECCIÓN CRÍTICA: Zona UTM 17S para Lambayeque
# Lambayeque: lon ~-80° → Zona 17, hemisferio Sur
# ============================================================
ZONE_NUMBER = 17
ZONE_LETTER = 'S'  # Sur (no M que es Norte)

def utm_to_latlon(este, norte):
    """
    Convierte UTM a Lat/Lon para Lambayeque, Perú.
    Zona 17, Hemisferio Sur (northern=False)
    """
    try:
        # Zona 17, hemisferio sur (northern=False)
        lat, lon = utm.to_latlon(
            float(este), 
            float(norte), 
            17, 
            northern=False  # ← CRÍTICO: Lambayeque está al sur del ecuador
        )
        return round(lat, 6), round(lon, 6)
    except Exception as e:
        print(f"Error conversión UTM: {e}, E={este}, N={norte}")
        return None, None

def google_maps_link(lat, lon):
    return f"https://www.google.com/maps?q={lat},{lon}&z=18" if lat and lon else ""

def whatsapp_link(phone, message):
    import urllib.parse
    return f"https://wa.me/{phone}?text={urllib.parse.quote(message)}"

def mailto_link(to, subject, body):
    import urllib.parse
    return f"mailto:{to}?subject={urllib.parse.quote(subject)}&body={urllib.parse.quote(body)}"

# ============================================================
# CARGA DEL EXCEL
# ============================================================
def cargar_excel(archivo):
    """Carga Excel con detección robusta de hoja y encabezados."""
    import re
    
    todas_hojas = pd.read_excel(archivo, sheet_name=None, header=None)
    
    # Seleccionar hoja con más datos
    mejor_hoja = None
    max_filas = 0
    for nombre_hoja, df_temp in todas_hojas.items():
        filas_validas = df_temp.dropna(how='all').shape[0]
        if filas_validas > max_filas:
            max_filas = filas_validas
            mejor_hoja = nombre_hoja
    
    print(f"✅ Hoja seleccionada: '{mejor_hoja}' ({max_filas} filas)")
    
    df = pd.read_excel(archivo, sheet_name=mejor_hoja, header=None)
    
    # Buscar fila de encabezados
    fila_encabezado = None
    for i in range(min(10, len(df))):
        fila_str = ' '.join([str(x) for x in df.iloc[i] if pd.notna(x)]).lower()
        if 'código alimentador' in fila_str or 'empresa' in fila_str:
            fila_encabezado = i
            break
    
    if fila_encabezado is None:
        fila_encabezado = 2
    
    print(f"✅ Fila de encabezados: {fila_encabezado}")
    
    df.columns = df.iloc[fila_encabezado]
    df = df.iloc[fila_encabezado + 1:].copy()
    df = df[~df.iloc[:, 0].astype(str).str.contains('---', na=False)]
    df = df.dropna(how='all').reset_index(drop=True)
    df.columns = [str(c).strip() for c in df.columns]
    
    # Renombrar columnas
    rename_map = {
        'Código': 'CODIGO',
        'Código Alimentador MT': 'AMT',
        'Observaciones': 'OBSERVACION',
        'Estado': 'ESTADO',
        'Distrito': 'DISTRITO',
        'UNIDAD': 'UNIDAD',
        'Tipo de Deficiencia': 'TIPO_DEFICIENCIA',
        'Región': 'REGION',
        'Provincia': 'PROVINCIA',
        'Prioridad': 'PRIORIDAD',
        'Tipo de Instalación': 'TIPO_INSTALACION',
        'Código Centro Transformación': 'COD_CT',
        'Estado Enero 2023': 'ESTADO_ENE2023',
        'UTM Este': 'UTM_E_ORIG',
        'UTM Norte': 'UTM_N_ORIG',
        'UTM E Ene2023': 'UTM_E',
        'UTM N Ene2023': 'UTM_N',
    }
    for k, v in rename_map.items():
        if k in df.columns and v not in df.columns:
            df[v] = df[k]
    
    if 'AMT' not in df.columns:
        for c in df.columns:
            if 'alimentador' in str(c).lower():
                df['AMT'] = df[c]
                break
    
    # Coordenadas UTM (preferir Ene2023)
    col_e = 'UTM_E' if 'UTM_E' in df.columns else 'UTM_E_ORIG'
    col_n = 'UTM_N' if 'UTM_N' in df.columns else 'UTM_N_ORIG'
    
    if col_e in df.columns and col_n in df.columns:
        coords = df.apply(lambda r: utm_to_latlon(r.get(col_e), r.get(col_n)), axis=1)
        df["LATITUD"] = [c[0] for c in coords]
        df["LONGITUD"] = [c[1] for c in coords]
    else:
        df["LATITUD"] = None
        df["LONGITUD"] = None
    
    df["MAPS_LINK"] = df.apply(lambda r: google_maps_link(r["LATITUD"], r["LONGITUD"]), axis=1)
    
    print(f"✅ Columnas finales: {list(df.columns)}")
    print(f"✅ AMT encontrado: {'AMT' in df.columns}")
    print(f"✅ Total filas: {len(df)}")
    
    # Verificar coordenadas
    if 'LATITUD' in df.columns:
        lat_validas = df['LATITUD'].dropna()
        if not lat_validas.empty:
            print(f"✅ Latitud promedio: {lat_validas.mean():.4f} (debe ser ~-6.7 para Lambayeque)")
            print(f"✅ Longitud promedio: {df['LONGITUD'].dropna().mean():.4f} (debe ser ~-80.0)")
    
    return df

# ============================================================
# RESUMEN
# ============================================================
def generar_resumen(df):
    total = len(df)
    por_estado = df["ESTADO"].value_counts().to_dict() if "ESTADO" in df else {}
    pend = sum(v for k, v in por_estado.items() if "Subsanar" in str(k) or "Pendiente" in str(k))
    return {
        "total": total,
        "pendientes": pend,
        "ejecutadas": total - pend,
        "por_estado": por_estado,
        "por_amtdf": df["AMT"].value_counts().to_dict() if "AMT" in df else {},
        "por_distrito": df["DISTRITO"].value_counts().to_dict() if "DISTRITO" in df else {},
    }

# ============================================================
# MAPA
# ============================================================
def _divicon(hexcolor):
    html = f"""
    <div style="background:{hexcolor};width:26px;height:26px;border-radius:50% 50% 50% 0;
                transform:rotate(-45deg);border:2px solid #fff;
                box-shadow:0 2px 6px rgba(0,59,112,.45);display:flex;align-items:center;
                justify-content:center;">
      <span style="transform:rotate(45deg);color:#fff;font-size:12px;">&#9889;</span>
    </div>"""
    return folium.DivIcon(html=html, icon_size=(26, 26), icon_anchor=(13, 26))

def crear_mapa(df_filtrado):
    d = df_filtrado.dropna(subset=["LATITUD", "LONGITUD"])
    if d.empty: return None

    m = folium.Map(location=[d["LATITUD"].mean(), d["LONGITUD"].mean()],
                   zoom_start=13, tiles="OpenStreetMap")

    for _, r in d.iterrows():
        pendiente = "Subsanar" in str(r.get("ESTADO", "")) or "Pendiente" in str(r.get("ESTADO", ""))
        color = BRAND["secondary_blue"] if pendiente else BRAND["brand_green"]
        popup = f"""
        <div style="font-family:Open Sans,Arial;color:#2C4D6F;font-size:12px;max-width:300px;">
          <div style="background:#003B70;color:#fff;padding:6px 10px;
                      font-family:Montserrat;font-weight:700;margin:-10px -10px 8px;">
            {r.get('CODIGO','')}
          </div>
          <b style="color:#003B70">AMT:</b> {r.get('AMT','')}<br>
          <b style="color:#003B70">Observación:</b> {str(r.get('OBSERVACION',''))[:120]}<br>
          <b style="color:#003B70">Estado:</b> {r.get('ESTADO','')}<br>
          <b style="color:#003B70">UTM 17S:</b> E={r.get('UTM_E','')} N={r.get('UTM_N','')}<br>
          <a href="{r.get('MAPS_LINK','')}" target="_blank"
             style="color:#0072CE;font-weight:600;">📍 Abrir en Google Maps</a>
        </div>"""
        folium.Marker([r["LATITUD"], r["LONGITUD"]],
                      popup=folium.Popup(popup, max_width=320),
                      tooltip=f"{r.get('CODIGO','')} · {r.get('AMT','')}",
                      icon=_divicon(color)).add_to(m)
    return m

# ============================================================
# PDF CON FOTOS
# ============================================================
class PDFReport(FPDF):
    def header(self):
        self.set_fill_color(*hex_rgb(BRAND["secondary_blue"]))
        self.rect(0, 0, 105, 10, "F")
        self.set_fill_color(*hex_rgb(BRAND["brand_green"]))
        self.rect(105, 0, 105, 10, "F")
        self.set_y(13)
        self.set_font("Arial", "B", 13)
        self.set_text_color(*hex_rgb(BRAND["primary_dark_blue"]))
        self.cell(0, 8, "ENSA | Informe de Observaciones MT", 0, 1, "C")
        self.set_font("Arial", "I", 8)
        self.set_text_color(*hex_rgb(BRAND["text_secondary"]))
        self.cell(0, 5, f"Trabajamos para tu Desarrollo | {datetime.now().strftime('%d/%m/%Y %H:%M')}", 0, 1, "C")
        self.ln(3)

    def footer(self):
        self.set_y(-14)
        self.set_draw_color(*hex_rgb(BRAND["primary_dark_blue"]))
        self.set_line_width(0.7)
        self.line(10, self.get_y(), 200, self.get_y())
        self.set_font("Arial", "I", 8)
        self.set_text_color(*hex_rgb(BRAND["text_secondary"]))
        self.cell(0, 8, f"ENSA · Página {self.page_no()}", 0, 0, "C")

    def titulo_seccion(self, txt):
        self.set_font("Arial", "B", 12)
        self.set_text_color(*hex_rgb(BRAND["primary_dark_blue"]))
        self.cell(0, 8, txt.upper(), 0, 1)
        self.set_draw_color(*hex_rgb(BRAND["brand_green"]))
        self.set_line_width(0.8)
        self.line(10, self.get_y(), 80, self.get_y())
        self.ln(3)

    def cuerpo(self, txt):
        self.set_font("Arial", "", 9.5)
        self.set_text_color(*hex_rgb(BRAND["text_secondary"]))
        self.multi_cell(0, 5, txt)

def generar_pdf(df, resumen, fotos_dict=None):
    """
    fotos_dict: {codigo: bytes_de_imagen}
    """
    pdf = PDFReport()
    pdf.add_page()
    pdf.titulo_seccion("Resumen Ejecutivo")
    pdf.cuerpo(f"Total de observaciones: {resumen['total']}")
    pdf.cuerpo(f"Pendientes (Por Subsanar): {resumen['pendientes']}")
    pdf.cuerpo(f"Ejecutadas / Subsanadas: {resumen['ejecutadas']}")
    pct = round(resumen["ejecutadas"] / resumen["total"] * 100, 1) if resumen["total"] else 0
    pdf.cuerpo(f"Avance de subsanación: {pct} %")
    pdf.ln(4)

    pdf.titulo_seccion("Cantidad por Alimentador MT (AMT)")
    pdf.set_font("Arial", "B", 9)
    pdf.set_text_color(*hex_rgb(BRAND["background_light"]))
    pdf.set_fill_color(*hex_rgb(BRAND["primary_dark_blue"]))
    for h, w in zip(["ALIMENTADOR MT", "CANTIDAD"], [120, 60]):
        pdf.cell(w, 6, h, 0, 0, "C", True)
    pdf.ln()
    pdf.set_font("Arial", "", 9)
    for i, (amt, cant) in enumerate(sorted(resumen["por_amtdf"].items(), key=lambda x: -x[1])):
        fill = i % 2 == 0
        if fill: pdf.set_fill_color(244, 247, 251)
        pdf.set_text_color(*hex_rgb(BRAND["text_secondary"]))
        pdf.cell(120, 5.5, str(amt), 0, 0, "L", fill)
        pdf.set_text_color(*hex_rgb(BRAND["primary_dark_blue"]))
        pdf.set_font("Arial", "B", 9)
        pdf.cell(60, 5.5, str(cant), 0, 1, "C", fill)
        pdf.set_font("Arial", "", 9)
    pdf.ln(5)

    pdf.titulo_seccion("Detalle de Observaciones")
    headers = ["CODIGO", "AMT", "OBSERVACION", "ESTADO", "UBICACION"]
    widths = [26, 20, 74, 26, 44]
    pdf.set_font("Arial", "B", 8)
    pdf.set_text_color(*hex_rgb(BRAND["background_light"]))
    pdf.set_fill_color(*hex_rgb(BRAND["primary_dark_blue"]))
    for h, w in zip(headers, widths):
        pdf.cell(w, 6, h, 0, 0, "C", True)
    pdf.ln()
    pdf.set_font("Arial", "", 7.5)
    for _, r in df.iterrows():
        pdf.set_text_color(*hex_rgb(BRAND["text_secondary"]))
        vals = [str(r.get("CODIGO", ""))[:16], str(r.get("AMT", ""))[:10],
                str(r.get("OBSERVACION", ""))[:70], str(r.get("ESTADO", ""))[:18],
                f"{r.get('LATITUD','')}, {r.get('LONGITUD','')}"]
        for v, w in zip(vals, widths):
            pdf.cell(w, 5, v, 1, 0, "L")
        pdf.ln()
    
    # Agregar fotos al final del PDF
    if fotos_dict:
        pdf.add_page()
        pdf.titulo_seccion("Fotografías de Campo")
        for codigo, img_bytes in fotos_dict.items():
            try:
                # Guardar temporalmente para fpdf
                img_path = f"/tmp/foto_{codigo}.jpg"
                with open(img_path, "wb") as f:
                    f.write(img_bytes)
                pdf.set_font("Arial", "B", 10)
                pdf.set_text_color(*hex_rgb(BRAND["primary_dark_blue"]))
                pdf.cell(0, 8, f"Código: {codigo}", 0, 1)
                pdf.image(img_path, x=10, w=190)
                pdf.ln(5)
                Path(img_path).unlink(missing_ok=True)
            except Exception as e:
                print(f"Error agregando foto {codigo}: {e}")
    
    return bytes(pdf.output())

# ============================================================
# WORD CON FOTOS
# ============================================================
def _run(run, hexcolor, bold=False, size=10, font=FONT_BODY):
    r, g, b = hex_rgb(hexcolor)
    run.font.color.rgb = RGBColor(r, g, b)
    run.bold = bold
    run.font.size = Pt(size)
    run.font.name = font
    return run

def _shade(cell, hexcolor):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:fill"), hexcolor.lstrip("#"))
    tcPr.append(shd)

def generar_word(df, resumen, fotos_dict=None):
    doc = Document()
    st = doc.styles["Normal"]
    st.font.name = FONT_BODY
    st.font.size = Pt(10)
    r, g, b = hex_rgb(BRAND["text_secondary"])
    st.font.color.rgb = RGBColor(r, g, b)

    t = doc.add_table(rows=2, cols=1)
    _shade(t.rows[0].cells[0], BRAND["primary_dark_blue"])
    p = t.rows[0].cells[0].paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _run(p.add_run("ENSA · INFORME DE OBSERVACIONES MT"), BRAND["background_light"], True, 16, FONT_TITLE)
    _shade(t.rows[1].cells[0], BRAND["secondary_blue"])
    p2 = t.rows[1].cells[0].paragraphs[0]
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _run(p2.add_run("Trabajamos para tu Desarrollo"), BRAND["background_light"], False, 9, FONT_BODY)
    doc.add_paragraph()
    _run(doc.add_paragraph().add_run(f"Fecha: {datetime.now().strftime('%d/%m/%Y %H:%M')}"), BRAND["text_secondary"], False, 9)

    h = doc.add_paragraph()
    _run(h.add_run("RESUMEN EJECUTIVO"), BRAND["primary_dark_blue"], True, 14, FONT_TITLE)
    for label, val in [("Total de observaciones", resumen["total"]), ("Pendientes (Por Subsanar)", resumen["pendientes"]), ("Ejecutadas / Subsanadas", resumen["ejecutadas"])]:
        p = doc.add_paragraph()
        _run(p.add_run(f"{label}: "), BRAND["text_secondary"], True, 10)
        _run(p.add_run(f"{val}"), BRAND["primary_dark_blue"], True, 11)

    doc.add_paragraph()
    h = doc.add_paragraph()
    _run(h.add_run("CANTIDAD POR ALIMENTADOR MT (AMT)"), BRAND["primary_dark_blue"], True, 13, FONT_TITLE)
    tbl = doc.add_table(rows=1, cols=2)
    tbl.style = "Light Grid Accent 1"
    for i, txt in enumerate(["Alimentador MT", "Cantidad"]):
        _shade(tbl.rows[0].cells[i], BRAND["primary_dark_blue"])
        _run(tbl.rows[0].cells[i].paragraphs[0].add_run(txt), BRAND["background_light"], True, 10, FONT_TITLE)
    for amt, cant in sorted(resumen["por_amtdf"].items(), key=lambda x: -x[1]):
        row = tbl.add_row().cells
        _run(row[0].paragraphs[0].add_run(str(amt)), BRAND["text_secondary"], False, 10)
        _run(row[1].paragraphs[0].add_run(str(cant)), BRAND["primary_dark_blue"], True, 10)

    doc.add_page_break()
    h = doc.add_paragraph()
    _run(h.add_run("DETALLE DE OBSERVACIONES"), BRAND["primary_dark_blue"], True, 14, FONT_TITLE)

    for _, r in df.iterrows():
        p = doc.add_paragraph()
        _run(p.add_run(f"Código {r.get('CODIGO','')}"), BRAND["primary_dark_blue"], True, 11, FONT_TITLE)
        for lbl, val in [("AMT", r.get("AMT", "")), ("Observación", r.get("OBSERVACION", "")), ("Estado", r.get("ESTADO", "")), ("Distrito", r.get("DISTRITO", "")), ("UTM 17S", f"E={r.get('UTM_E','')} N={r.get('UTM_N','')}"), ("Lat/Lon", f"{r.get('LATITUD','')}, {r.get('LONGITUD','')}"), ("Mapa", r.get("MAPS_LINK", ""))]:
            pp = doc.add_paragraph()
            _run(pp.add_run(f"{lbl}: "), BRAND["primary_dark_blue"], True, 9.5)
            _run(pp.add_run(str(val)), BRAND["text_secondary"], False, 9.5)
        
        # Agregar foto si existe
        if fotos_dict and r.get('CODIGO') in fotos_dict:
            img_bytes = fotos_dict[r.get('CODIGO')]
            img_stream = io.BytesIO(img_bytes)
            doc.add_picture(img_stream, width=Inches(5))
            doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        doc.add_paragraph("_" * 90)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.getvalue()

# ============================================================
# EXPORTAR A EXCEL (no CSV)
# ============================================================
def exportar_a_excel(df, resumen):
    """Exporta el DataFrame a Excel nativo con formato."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Hoja 1: Datos completos
        df.to_excel(writer, sheet_name='Observaciones', index=False)
        
        # Hoja 2: Resumen
        resumen_df = pd.DataFrame([
            {"Métrica": "Total", "Valor": resumen["total"]},
            {"Métrica": "Pendientes", "Valor": resumen["pendientes"]},
            {"Métrica": "Ejecutadas", "Valor": resumen["ejecutadas"]},
        ])
        resumen_df.to_excel(writer, sheet_name='Resumen', index=False)
        
        # Hoja 3: Por AMT
        amt_df = pd.DataFrame([{"AMT": k, "Cantidad": v} 
                               for k, v in sorted(resumen["por_amtdf"].items(), key=lambda x: -x[1])])
        amt_df.to_excel(writer, sheet_name='Por_AMT', index=False)
        
        # Hoja 4: Por Distrito
        dis_df = pd.DataFrame([{"Distrito": k, "Cantidad": v} 
                               for k, v in sorted(resumen["por_distrito"].items(), key=lambda x: -x[1])])
        dis_df.to_excel(writer, sheet_name='Por_Distrito', index=False)
    
    output.seek(0)
    return output.getvalue()