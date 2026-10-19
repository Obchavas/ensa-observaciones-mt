# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import altair as alt
from datetime import datetime
from pathlib import Path

from utils import (BRAND, cargar_excel, generar_resumen, crear_mapa,
                   generar_pdf, generar_word, whatsapp_link, mailto_link,
                   exportar_a_excel)

st.set_page_config(page_title="ENSA · Observaciones MT", page_icon="⚡", layout="wide", initial_sidebar_state="expanded")

# ============================================================
# CSS CORPORATIVO
# ============================================================
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@600;700;800&family=Open+Sans:wght@400;600&display=swap');

* { font-family: 'Open Sans', Arial, sans-serif !important; }
.stApp { background: #FFFFFF !important; }
h1, h2, h3, h4 { color: #003B70 !important; font-family: 'Montserrat', Arial, sans-serif !important; font-weight: 800 !important; }

[data-testid="stSidebar"] { background: #FFFFFF !important; border-right: 3px solid #0072CE !important; }
[data-testid="stSidebar"] h1, [data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3 { color: #003B70 !important; }

.stButton > button {
    background: #0072CE !important; color: #FFFFFF !important; border: none !important;
    border-radius: 10px !important; font-family: 'Montserrat', Arial, sans-serif !important;
    font-weight: 700 !important; padding: 10px 20px !important;
}
.stButton > button:hover { background: #003B70 !important; color: #FFFFFF !important; }

.stDownloadButton > button {
    background: #38B449 !important; color: #FFFFFF !important; border: none !important;
    border-radius: 10px !important; font-weight: 700 !important;
}
.stDownloadButton > button:hover { background: #003B70 !important; }

.stTabs [data-baseweb="tab-list"] { gap: 8px; }
.stTabs [data-baseweb="tab"] {
    font-family: 'Montserrat', Arial, sans-serif !important; font-weight: 700 !important;
    color: #2C4D6F !important; background: #F4F7FB !important; border-radius: 8px 8px 0 0 !important;
}
.stTabs [aria-selected="true"] {
    color: #FFFFFF !important; background: #003B70 !important; border-bottom: 3px solid #38B449 !important;
}

.stAlert { border-left: 5px solid #0072CE !important; background: #F4F7FB !important; }
thead th { background: #003B70 !important; color: #FFFFFF !important; font-family: 'Montserrat' !important; }

.dl-banner {
    background: linear-gradient(90deg, #0072CE, #38B449); color: #fff;
    padding: 24px 30px; border-radius: 14px; margin-bottom: 20px;
    box-shadow: 0 6px 18px rgba(0,59,112,.18);
}
.dl-banner h1 { color: #fff !important; margin: 0; font-size: 1.6rem; }
.dl-banner p { color: #fff; margin: 4px 0 0; font-style: italic; font-size: 0.95rem; }

.dl-footer {
    background: #003B70; color: #fff; text-align: center;
    padding: 14px; border-radius: 12px; margin-top: 30px; font-size: 0.85rem;
}

.dl-card {
    position: relative; background: #FFFFFF; border: 1px solid #E3EAF2;
    border-radius: 14px; padding: 22px 24px; box-shadow: 0 3px 12px rgba(0,59,112,.08);
    overflow: hidden; margin-bottom: 10px;
}
.dl-card::before {
    content: ""; position: absolute; top: 0; left: 0; right: 0; height: 5px;
    background: linear-gradient(90deg, #0072CE, #38B449);
}
.dl-card .lbl {
    font-family: 'Montserrat', Arial, sans-serif; font-size: 0.72rem;
    letter-spacing: 0.08em; text-transform: uppercase; color: #2C4D6F; font-weight: 600;
}
.dl-card .num {
    font-family: 'Montserrat', Arial, sans-serif; font-size: 2.8rem;
    font-weight: 800; color: #003B70; line-height: 1.1; margin: 6px 0;
}
.dl-card .sub { font-size: 0.8rem; color: #2C4D6F; }

.dl-bar { background: #F4F7FB; border-radius: 20px; height: 14px; overflow: hidden; margin-top: 10px; }
.dl-bar > div { height: 100%; background: linear-gradient(90deg, #0072CE, #38B449); border-radius: 20px; }
</style>
""", unsafe_allow_html=True)

# ============================================================
# BARRA SUPERIOR
# ============================================================
logo_path = Path("logo_ensa.png")
logo_url = None
if logo_path.exists():
    import base64
    with open(logo_path, "rb") as f:
        logo_url = "data:image/png;base64," + base64.b64encode(f.read()).decode()

top1, top2 = st.columns([1, 4])
with top1:
    if logo_url:
        st.image(logo_url, width=130)
    else:
        st.markdown("###  ENSA")
with top2:
    st.markdown("### Sistema de Observaciones MT · Meta 2022")
    st.caption("*Trabajamos para tu Desarrollo · Zona UTM 17S Lambayeque*")

st.markdown("---")

# ============================================================
# SIDEBAR
# ============================================================
st.sidebar.subheader("📁 Cargar Archivo")
archivo = st.sidebar.file_uploader("Excel (MetaMT_2022.xlsx)", type=["xlsx", "xls"])

if archivo is None:
    st.markdown("""<div class="dl-banner">
    <h1>⚡ Informe de Observaciones MT · MetaMT 2022</h1>
    <p>Verificación en campo con coordenadas UTM 17S / geográficas y mapa interactivo</p>
    </div>""", unsafe_allow_html=True)
    st.info("👈 Sube el archivo Excel desde el panel izquierdo para comenzar.")
    st.stop()

with st.spinner("Procesando archivo..."):
    try:
        df = cargar_excel(archivo)
    except Exception as e:
        st.error(f" Error al procesar: {e}")
        st.stop()

if df.empty:
    st.error("El archivo no contiene datos válidos.")
    st.stop()

if 'AMT' not in df.columns:
    st.error(f"❌ No se encontró 'Código Alimentador MT'. Columnas: {list(df.columns)}")
    st.stop()

# Filtros
st.sidebar.markdown("---")
st.sidebar.subheader(" Filtros")

amts = sorted([str(x) for x in df["AMT"].dropna().unique().tolist()])
amt_sel = st.sidebar.multiselect("Alimentador MT (AMT)", amts, default=amts[:3] if len(amts) > 3 else amts)

estados = df["ESTADO"].dropna().unique().tolist() if "ESTADO" in df.columns else []
est_sel = st.sidebar.multiselect("Estado", estados, default=estados)

distritos = df["DISTRITO"].dropna().unique().tolist() if "DISTRITO" in df.columns else []
dis_sel = st.sidebar.multiselect("Distrito", distritos)

busq = st.sidebar.text_input(" Buscar en observación", "")

d = df.copy()
if amt_sel:  d = d[d["AMT"].astype(str).isin(amt_sel)]
if est_sel and "ESTADO" in d.columns:  d = d[d["ESTADO"].isin(est_sel)]
if dis_sel and "DISTRITO" in d.columns:  d = d[d["DISTRITO"].isin(dis_sel)]
if busq and "OBSERVACION" in d.columns:
    d = d[d["OBSERVACION"].astype(str).str.contains(busq, case=False, na=False)]

st.sidebar.success(f"✅ {len(d)} registros filtrados")

# ============================================================
# DASHBOARD
# ============================================================
res = generar_resumen(d)
pct = round(res["ejecutadas"] / res["total"] * 100, 1) if res["total"] else 0

st.markdown("""<div class="dl-banner">
<h1>📊 Dashboard de Observaciones MT</h1>
<p>Resumen ejecutivo del estado de las instalaciones</p>
</div>""", unsafe_allow_html=True)

c1, c2, c3, c4 = st.columns(4)

def card(col, lbl, num, sub, w):
    with col:
        st.markdown(f"""<div class="dl-card">
        <div class="lbl">{lbl}</div>
        <div class="num">{num}</div>
        <div class="sub">{sub}</div>
        <div class="dl-bar"><div style="width:{min(w,100):.0f}%"></div></div>
        </div>""", unsafe_allow_html=True)

card(c1, "Total", res["total"], "Observaciones filtradas", 100)
card(c2, "Pendientes", res["pendientes"], "Por Subsanar", res["pendientes"]/res["total"]*100 if res["total"] else 0)
card(c3, "Ejecutadas", res["ejecutadas"], "Subsanadas en campo", res["ejecutadas"]/res["total"]*100 if res["total"] else 0)
card(c4, "Avance", f"{pct}%", "Cumplimiento de la meta", pct)

st.markdown("---")

# ============================================================
# TABS
# ============================================================
tab1, tab2, tab3, tab4, tab5 = st.tabs(["🗺️ Mapa", "📋 Listado", "📊 Resumen AMT", "📸 Captura en Campo", "📥 Reportes"])

# ----- MAPA -----
with tab1:
    st.subheader("Mapa de observaciones georreferenciadas (UTM 17S Lambayeque)")
    mapa = crear_mapa(d)
    if mapa:
        from streamlit_folium import st_folium
        st_folium(mapa, width=None, height=620, returned_objects=[])
    else:
        st.warning("Sin coordenadas válidas para graficar.")

# ----- LISTADO -----
with tab2:
    st.subheader("Listado de observaciones")
    cols = [c for c in ["CODIGO", "AMT", "OBSERVACION", "ESTADO", "DISTRITO", "LATITUD", "LONGITUD", "MAPS_LINK"] if c in d.columns]
    tbl = d[cols].copy()
    if "MAPS_LINK" in tbl.columns:
        tbl["MAPS_LINK"] = tbl["MAPS_LINK"].apply(lambda x: f'<a href="{x}" target="_blank">📍 Ver mapa</a>' if x else "")
    st.dataframe(tbl, use_container_width=True, height=520)

# ----- RESUMEN AMT -----
with tab3:
    st.subheader("Cantidad de observaciones por Alimentador MT (AMT)")
    if res["por_amtdf"]:
        damt = pd.DataFrame([{"AMT": str(k), "Cantidad": v} for k, v in sorted(res["por_amtdf"].items(), key=lambda x: -x[1])])
        ch = alt.Chart(damt).mark_bar(cornerRadiusTopLeft=6, cornerRadiusTopRight=6).encode(
            x=alt.X("AMT:N", sort="-y"), y=alt.Y("Cantidad:Q"), color=alt.value(BRAND["secondary_blue"])
        ).properties(height=380)
        st.altair_chart(ch, use_container_width=True)
        st.dataframe(damt, use_container_width=True)

# ----- CAPTURA EN CAMPO CON FOTOS -----
with tab4:
    st.subheader(" Captura de observación en campo")
    st.info("Selecciona el código, toma la foto con GPS activo y registra el estado real.")
    
    cods = d["CODIGO"].dropna().unique().tolist()
    cod = st.selectbox("Código a verificar", cods)
    
    if cod:
        r = d[d["CODIGO"] == cod].iloc[0]
        a, b = st.columns(2)
        with a:
            st.markdown(f"**Observación:** {r.get('OBSERVACION','')}")
            st.markdown(f"**AMT:** {r.get('AMT','')} · **Distrito:** {r.get('DISTRITO','')}")
            st.markdown(f"**Estado:** {r.get('ESTADO','')}")
        with b:
            st.markdown(f"**UTM 17S:** E={r.get('UTM_E','')} · N={r.get('UTM_N','')}")
            st.markdown(f"**Geográficas:** {r.get('LATITUD','')} , {r.get('LONGITUD','')}")
            if r.get("MAPS_LINK"):
                st.markdown(f"[ Abrir Google Maps]({r['MAPS_LINK']})")

        foto = st.camera_input(" Tomar foto con coordenadas GPS")
        nuevo = st.selectbox("Estado tras la verificación", ["Pendiente", "Ejecutada", "Requiere replanificación"])
        obs = st.text_area("Observaciones de campo", "")

        if foto is not None:
            st.image(foto, caption="Foto capturada en campo", use_column_width=True)
            
            # Guardar foto en session_state
            if "fotos_campo" not in st.session_state:
                st.session_state.fotos_campo = {}
            st.session_state.fotos_campo[cod] = foto.getvalue()
            
            st.success(f"✅ Foto guardada para código {cod}")
            
            msg = (f"*ENSA · VERIFICACION EN CAMPO*\n"
                   f"Codigo: {cod}\nAMT: {r.get('AMT','')}\nEstado: {nuevo}\n"
                   f"UTM 17S: E={r.get('UTM_E','')} N={r.get('UTM_N','')}\n"
                   f"Lat/Lon: {r.get('LATITUD','')}, {r.get('LONGITUD','')}\n"
                   f"Mapa: {r.get('MAPS_LINK','')}\nObs: {obs}")
            tel = st.text_input("WhatsApp (con código país, ej 51987654321)", "51")
            cA, cB = st.columns(2)
            with cA:
                if st.button("📲 Enviar por WhatsApp"):
                    st.markdown(f"[Abrir WhatsApp]({whatsapp_link(tel, msg)})")
            with cB:
                st.download_button("💾 Descargar foto", foto.getvalue(),
                                   file_name=f"obs_{cod}_{datetime.now():%Y%m%d_%H%M%S}.jpg",
                                   mime="image/jpeg")

# ----- REPORTES -----
with tab5:
    st.subheader("Generación y envío de reportes")
    
    # Mostrar fotos capturadas
    if "fotos_campo" in st.session_state and st.session_state.fotos_campo:
        st.info(f"📸 {len(st.session_state.fotos_campo)} foto(s) capturada(s) en campo")
    
    r1, r2, r3 = st.columns(3)
    
    with r1:
        st.markdown("#### 📄 Reporte PDF")
        if st.button("Generar PDF", use_container_width=True):
            fotos = st.session_state.get("fotos_campo", {})
            st.download_button("⬇️ Descargar PDF", data=generar_pdf(d, res, fotos),
                               file_name=f"ENSA_Obs_MT_{datetime.now():%Y%m%d}.pdf", mime="application/pdf", use_container_width=True)
    
    with r2:
        st.markdown("#### 📝 Reporte Word")
        if st.button("Generar Word", use_container_width=True):
            fotos = st.session_state.get("fotos_campo", {})
            st.download_button("⬇️ Descargar Word", data=generar_word(d, res, fotos),
                               file_name=f"ENSA_Obs_MT_{datetime.now():%Y%m%d}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
    
    with r3:
        st.markdown("#### 📊 Reporte Excel")
        if st.button("Generar Excel", use_container_width=True):
            st.download_button("⬇️ Descargar Excel", data=exportar_a_excel(d, res),
                               file_name=f"ENSA_Obs_MT_{datetime.now():%Y%m%d}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

    st.markdown("---")
    st.markdown("#### 📤 Envío rápido")
    resumen_txt = (f"*ENSA · RESUMEN OBSERVACIONES MT*\n"
                   f"Fecha: {datetime.now():%d/%m/%Y}\n"
                   f"Total: {res['total']} | Pendientes: {res['pendientes']} | "
                   f"Ejecutadas: {res['ejecutadas']} | Avance: {pct}%\n\n*Top AMT:*\n")
    for i, (k, v) in enumerate(list(res["por_amtdf"].items())[:5], 1):
        resumen_txt += f"{i}. {k}: {v}\n"

    cc1, cc2 = st.columns(2)
    with cc1:
        tel = st.text_input("WhatsApp destino", "51", key="tel2")
        if st.button("📲 Enviar resumen por WhatsApp"):
            st.markdown(f"[Abrir WhatsApp con el resumen]({whatsapp_link(tel, resumen_txt)})")
    with cc2:
        mail = st.text_input("Correo destino", "supervisor@ensa.com.pe")
        if st.button("📧 Preparar correo"):
            st.markdown(f"[Abrir correo con el resumen]({mailto_link(mail, 'ENSA - Resumen Observaciones MT', resumen_txt)})")

# ============================================================
# PIE
# ============================================================
st.markdown(f"""<div class="dl-footer">
<b>ENSA</b> · Trabajamos para tu Desarrollo ·
Área de Distribución MT · Documento Osinergmin MetaMT 2022 ·
Zona UTM 17S Lambayeque · {datetime.now():%d/%m/%Y %H:%M}
</div>""", unsafe_allow_html=True)