# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import altair as alt
from datetime import datetime
from pathlib import Path

from utils import (BRAND, cargar_excel, generar_resumen, crear_mapa,
                   generar_pdf, generar_word, whatsapp_link, mailto_link,
                   exportar_a_excel, exportar_fotos_zip)

st.set_page_config(page_title="ENSA · Observaciones MT", page_icon="", layout="wide", initial_sidebar_state="collapsed")

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

.avance-verde { background: linear-gradient(90deg, #38B449, #0072CE) !important; }
.avance-parcial { background: linear-gradient(90deg, #FFA500, #0072CE) !important; }
</style>
""", unsafe_allow_html=True)

# ============================================================
# INICIALIZAR SESSION STATE
# ============================================================
if "datos_cargados" not in st.session_state:
    st.session_state.datos_cargados = False
if "df_completo" not in st.session_state:
    st.session_state.df_completo = None
if "avance_por_codigo" not in st.session_state:
    st.session_state.avance_por_codigo = {}

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
        st.markdown("### ⚡ ENSA")
with top2:
    st.markdown("### Sistema de Observaciones MT · Meta 2022")
    st.caption("*Trabajamos para tu Desarrollo · Zona UTM 17S Lambayeque*")

st.markdown("---")

# ============================================================
# TABS PRINCIPALES
# ============================================================
tab_admin, tab_campo = st.tabs(["👨‍💼 Administrador (Carga de Datos)", "👷 Técnico de Campo"])

# ============================================================
# PESTAÑA ADMINISTRADOR
# ============================================================
with tab_admin:
    st.markdown("### 👨‍💼 Panel de Administrador")
    st.info("🔒 **Solo tú tienes acceso a esta pestaña** para cargar el Excel maestro")
    
    if not st.session_state.datos_cargados:
        archivo = st.file_uploader(" Cargar Excel Maestro (MetaMT_2022.xlsx)", type=["xlsx", "xls"])
        
        if archivo is not None:
            with st.spinner("Procesando archivo..."):
                try:
                    df = cargar_excel(archivo)
                    st.session_state.df_completo = df
                    st.session_state.datos_cargados = True
                    st.success(f"✅ {len(df)} observaciones cargadas correctamente")
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ Error: {e}")
    else:
        st.success("✅ Excel cargado correctamente")
        st.write(f"**Total de observaciones:** {len(st.session_state.df_completo)}")
        
        # Mostrar resumen por alimentador
        if "AMT" in st.session_state.df_completo.columns:
            st.markdown("### 📊 Resumen por Alimentador")
            resumen_amt = st.session_state.df_completo["AMT"].value_counts().reset_index()
            resumen_amt.columns = ["Alimentador", "Total Observaciones"]
            st.dataframe(resumen_amt, use_container_width=True)
        
        # Botón para descargar base actualizada
        if st.button("📥 Descargar Base de Datos Actualizada"):
            df_export = st.session_state.df_completo.copy()
            df_export["VISITADO"] = df_export["CODIGO"].map(
                lambda x: st.session_state.avance_por_codigo.get(x, {}).get("visitado", False)
            )
            df_export["OBS_CAMPO"] = df_export["CODIGO"].map(
                lambda x: st.session_state.avance_por_codigo.get(x, {}).get("obs", "")
            )
            
            excel_data = exportar_a_excel(df_export, generar_resumen(df_export))
            st.download_button(
                "️ Descargar Excel con Avance",
                data=excel_data,
                file_name=f"ENSA_Obs_MT_Actualizado_{datetime.now():%Y%m%d}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        
        # Botón para reiniciar
        if st.button("🔄 Cargar Nuevo Excel"):
            st.session_state.datos_cargados = False
            st.session_state.df_completo = None
            st.session_state.avance_por_codigo = {}
            st.rerun()

# ============================================================
# PESTAÑA TÉCNICO DE CAMPO (COMPLETA CON REPORTES)
# ============================================================
with tab_campo:
    if not st.session_state.datos_cargados:
        st.warning("️ El administrador aún no ha cargado los datos. Contacta con tu supervisor.")
        st.stop()
    
    st.markdown("### 👷 Panel de Técnico de Campo")
    
    # Selector de alimentador
    if "AMT" in st.session_state.df_completo.columns:
        amts_disponibles = sorted(st.session_state.df_completo["AMT"].dropna().unique().tolist())
        
        col1, col2 = st.columns([3, 1])
        with col1:
            amt_seleccionado = st.selectbox(
                "🔌 Selecciona tu Alimentador MT asignado:",
                amts_disponibles,
                index=0 if amts_disponibles else None
            )
        
        # Filtrar datos por alimentador
        df_filtrado = st.session_state.df_completo[
            st.session_state.df_completo["AMT"] == amt_seleccionado
        ].copy()
        
        # Calcular avance
        total_obs = len(df_filtrado)
        visitados = sum(1 for cod in df_filtrado["CODIGO"] if st.session_state.avance_por_codigo.get(cod, {}).get("visitado", False))
        porcentaje = (visitados / total_obs * 100) if total_obs > 0 else 0
        
        with col2:
            st.metric("📈 Avance", f"{porcentaje:.1f}%", f"{visitados}/{total_obs}")
        
        st.markdown("---")
        
        # Dashboard de avance
        st.markdown("###  Tu Avance - " + amt_seleccionado)
        c1, c2, c3, c4 = st.columns(4)
        
        with c1:
            st.markdown(f"""<div class="dl-card">
            <div class="lbl">Total Observaciones</div>
            <div class="num">{total_obs}</div>
            <div class="sub">Asignadas a este AMT</div>
            </div>""", unsafe_allow_html=True)
        
        with c2:
            st.markdown(f"""<div class="dl-card">
            <div class="lbl">Visitadas</div>
            <div class="num">{visitados}</div>
            <div class="sub">Completadas en campo</div>
            </div>""", unsafe_allow_html=True)
        
        with c3:
            st.markdown(f"""<div class="dl-card">
            <div class="lbl">Pendientes</div>
            <div class="num">{total_obs - visitados}</div>
            <div class="sub">Por visitar</div>
            </div>""", unsafe_allow_html=True)
        
        with c4:
            clase_barra = "avance-verde" if porcentaje == 100 else "avance-parcial" if porcentaje > 0 else ""
            st.markdown(f"""<div class="dl-card">
            <div class="lbl">Porcentaje</div>
            <div class="num">{porcentaje:.0f}%</div>
            <div class="dl-bar"><div class="{clase_barra}" style="width:{porcentaje}%"></div></div>
            </div>""", unsafe_allow_html=True)
        
        st.markdown("---")
        
        # Tabs internos para Técnico de Campo
        tab_ver, tab_registrar, tab_reportes = st.tabs(["📋 Ver Observaciones", "📸 Registrar en Campo", "📥 Generar Reportes"])
        
        # ----- TAB: VER OBSERVACIONES -----
        with tab_ver:
            st.markdown("### 📋 Listado de Observaciones")
            
            # Filtros adicionales
            col_f1, col_f2 = st.columns(2)
            with col_f1:
                estado_filtro = st.multiselect(
                    "Filtrar por Estado:",
                    options=df_filtrado["ESTADO"].dropna().unique().tolist() if "ESTADO" in df_filtrado.columns else [],
                    default=[]
                )
            with col_f2:
                distrito_filtro = st.multiselect(
                    "Filtrar por Distrito:",
                    options=df_filtrado["DISTRITO"].dropna().unique().tolist() if "DISTRITO" in df_filtrado.columns else [],
                    default=[]
                )
            
            # Aplicar filtros
            df_mostrar = df_filtrado.copy()
            if estado_filtro:
                df_mostrar = df_mostrar[df_mostrar["ESTADO"].isin(estado_filtro)]
            if distrito_filtro:
                df_mostrar = df_mostrar[df_mostrar["DISTRITO"].isin(distrito_filtro)]
            
            # Mostrar tabla con indicador de visitado
            df_mostrar["VISITADA"] = df_mostrar["CODIGO"].map(
                lambda x: "✅ Visitada" if st.session_state.avance_por_codigo.get(x, {}).get("visitado", False) else "⏳ Pendiente"
            )
            
            cols_tabla = ["VISITADA", "CODIGO", "OBSERVACION", "ESTADO", "DISTRITO", "LATITUD", "LONGITUD"]
            cols_disponibles = [c for c in cols_tabla if c in df_mostrar.columns]
            
            st.dataframe(
                df_mostrar[cols_disponibles],
                use_container_width=True,
                height=400
            )
            
            # Mapa
            st.markdown("---")
            st.markdown("### 🗺️ Mapa de Observaciones")
            mapa = crear_mapa(df_filtrado)
            if mapa:
                from streamlit_folium import st_folium
                st_folium(mapa, width=None, height=500, returned_objects=[])
            else:
                st.warning("Sin coordenadas válidas para graficar.")
        
        # ----- TAB: REGISTRAR EN CAMPO -----
        with tab_registrar:
            st.markdown("### 📸 Registrar Visita en Campo")
            
            codigos_pendientes = [
                cod for cod in df_filtrado["CODIGO"] 
                if not st.session_state.avance_por_codigo.get(cod, {}).get("visitado", False)
            ]
            
            if codigos_pendientes:
                codigo_seleccionado = st.selectbox(
                    "Selecciona el código a registrar:",
                    codigos_pendientes
                )
                
                # Mostrar datos del código seleccionado
                obs_data = df_filtrado[df_filtrado["CODIGO"] == codigo_seleccionado].iloc[0]
                
                col_info1, col_info2 = st.columns(2)
                with col_info1:
                    st.markdown(f"""
                    **Código:** {codigo_seleccionado}  
                    **AMT:** {obs_data.get('AMT', 'N/A')}  
                    **Observación:** {obs_data.get('OBSERVACION', 'N/A')}  
                    **Estado Original:** {obs_data.get('ESTADO', 'N/A')}
                    """)
                
                with col_info2:
                    st.markdown(f"""
                    **Distrito:** {obs_data.get('DISTRITO', 'N/A')}  
                    **UTM E:** {obs_data.get('UTM_E', 'N/A')}  
                    **UTM N:** {obs_data.get('UTM_N', 'N/A')}  
                    **Lat/Lon:** {obs_data.get('LATITUD', 'N/A')}, {obs_data.get('LONGITUD', 'N/A')}
                    """)
                
                # Formulario de registro
                with st.form(f"registro_{codigo_seleccionado}"):
                    foto = st.camera_input("📸 Tomar foto en campo")
                    if foto is None:
                        foto = st.file_uploader(" O subir foto desde galería", type=['jpg', 'jpeg', 'png'], key=f"upload_{codigo_seleccionado}")
                    
                    obs_campo = st.text_area("📝 Observaciones de campo (opcional):")
                    estado_final = st.selectbox(
                        "Estado después de la visita:",
                        ["Pendiente", "Ejecutada", "Requiere replanificación"]
                    )
                    
                    submitted = st.form_submit_button("💾 Guardar Visita", use_container_width=True)
                    
                    if submitted:
                        if foto is not None:
                            st.session_state.avance_por_codigo[codigo_seleccionado] = {
                                "visitado": True,
                                "foto": foto.getvalue() if hasattr(foto, 'getvalue') else foto.read(),
                                "obs": obs_campo,
                                "estado_final": estado_final,
                                "fecha_visita": datetime.now().isoformat()
                            }
                            st.success(f"✅ Visita registrada para {codigo_seleccionado}")
                            st.rerun()
                        else:
                            st.error("❌ Debes tomar o subir una foto")
            else:
                st.success("🎉 ¡Felicitaciones! Has visitado todas las observaciones de este alimentador.")
        
        # ----- TAB: GENERAR REPORTES (COMPLETO) -----
        with tab_reportes:
            st.markdown("###  Generar Reportes")
            
            fotos_count = len([cod for cod in df_filtrado["CODIGO"] if cod in st.session_state.avance_por_codigo])
            
            if fotos_count > 0:
                st.success(f"📸 {fotos_count} observación(es) registrada(s) con foto")
            else:
                st.info("No hay observaciones registradas aún. Ve a la pestaña ' Registrar en Campo'.")
            
            st.markdown("---")
            
            # Botones de reportes
            r1, r2, r3, r4 = st.columns(4)
            
            with r1:
                st.markdown("#### 📄 Reporte PDF")
                if st.button("Generar PDF", use_container_width=True):
                    # Filtrar solo los registros con foto
                    codigos_con_foto = [cod for cod in df_filtrado["CODIGO"] if cod in st.session_state.avance_por_codigo]
                    df_con_fotos = df_filtrado[df_filtrado["CODIGO"].isin(codigos_con_foto)]
                    fotos_dict = {cod: st.session_state.avance_por_codigo[cod]["foto"] for cod in codigos_con_foto}
                    
                    st.download_button(
                        "⬇️ Descargar PDF",
                        data=generar_pdf(df_con_fotos, generar_resumen(df_con_fotos), fotos_dict),
                        file_name=f"ENSA_{amt_seleccionado}_{datetime.now():%Y%m%d}.pdf",
                        mime="application/pdf",
                        use_container_width=True
                    )
            
            with r2:
                st.markdown("#### 📝 Reporte Word")
                if st.button("Generar Word", use_container_width=True):
                    codigos_con_foto = [cod for cod in df_filtrado["CODIGO"] if cod in st.session_state.avance_por_codigo]
                    df_con_fotos = df_filtrado[df_filtrado["CODIGO"].isin(codigos_con_foto)]
                    fotos_dict = {cod: st.session_state.avance_por_codigo[cod]["foto"] for cod in codigos_con_foto}
                    
                    st.download_button(
                        "⬇️ Descargar Word",
                        data=generar_word(df_con_fotos, generar_resumen(df_con_fotos), fotos_dict),
                        file_name=f"ENSA_{amt_seleccionado}_{datetime.now():%Y%m%d}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True
                    )
            
            with r3:
                st.markdown("#### 📊 Reporte Excel")
                if st.button("Generar Excel", use_container_width=True):
                    codigos_con_foto = [cod for cod in df_filtrado["CODIGO"] if cod in st.session_state.avance_por_codigo]
                    df_con_fotos = df_filtrado[df_filtrado["CODIGO"].isin(codigos_con_foto)]
                    
                    # Agregar columnas de avance
                    df_export = df_con_fotos.copy()
                    df_export["VISITADO"] = df_export["CODIGO"].map(
                        lambda x: st.session_state.avance_por_codigo.get(x, {}).get("visitado", False)
                    )
                    df_export["OBS_CAMPO"] = df_export["CODIGO"].map(
                        lambda x: st.session_state.avance_por_codigo.get(x, {}).get("obs", "")
                    )
                    
                    st.download_button(
                        "️ Descargar Excel",
                        data=exportar_a_excel(df_export, generar_resumen(df_con_fotos)),
                        file_name=f"ENSA_{amt_seleccionado}_{datetime.now():%Y%m%d}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
            
            with r4:
                st.markdown("#### 📦 Extraer Fotos ZIP")
                if st.button("Extraer Fotos", use_container_width=True):
                    codigos_con_foto = [cod for cod in df_filtrado["CODIGO"] if cod in st.session_state.avance_por_codigo]
                    fotos_dict = {cod: st.session_state.avance_por_codigo[cod]["foto"] for cod in codigos_con_foto}
                    
                    if fotos_dict:
                        zip_data = exportar_fotos_zip(fotos_dict, df_filtrado)
                        st.download_button(
                            "⬇️ Descargar ZIP",
                            data=zip_data,
                            file_name=f"Fotos_{amt_seleccionado}_{datetime.now():%Y%m%d}.zip",
                            mime="application/zip",
                            use_container_width=True
                        )
                    else:
                        st.warning("No hay fotos para extraer")

# ============================================================
# PIE
# ============================================================
st.markdown(f"""<div class="dl-footer">
<b>ENSA</b> · Trabajamos para tu Desarrollo ·
Área de Distribución MT · Documento Osinergmin MetaMT 2022 ·
Zona UTM 17S Lambayeque · {datetime.now():%d/%m/%Y %H:%M}
</div>""", unsafe_allow_html=True)