import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import matplotlib.pyplot as plt
from datetime import datetime
import json
import io
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

# ── Configuración de página ───────────────────────────────────────────────────
st.set_page_config(
    page_title="Sistema Académico · Fuente de Gracia",
    page_icon="✝️",
    layout="wide",
    initial_sidebar_state="expanded"
)

SHEET_ID = "1biSSyFbRPv3JNCxaGoHzewxVEd7S80tm4yGAbA9jpEo"
SCOPES   = ["https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"]

# ── CSS Institucional ─────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Playfair+Display:wght@400;600;700&family=Source+Sans+3:wght@300;400;600&display=swap');

html, body, [class*="css"] { font-family: 'Source Sans 3', sans-serif; }

.inst-header {
    background: linear-gradient(135deg, #0d1b3e 0%, #1a2f5e 60%, #0d1b3e 100%);
    border-bottom: 3px solid #c9a84c;
    padding: 22px 32px;
    margin: -1rem -1rem 2rem -1rem;
    display: flex; align-items: center; gap: 20px;
}
.inst-header h1 {
    font-family: 'Playfair Display', serif;
    color: #c9a84c; font-size: 1.5rem; margin: 0; line-height: 1.2;
}
.inst-header p {
    color: #a0aec0; font-size: .8rem; margin: 2px 0 0;
    letter-spacing: 1.5px; text-transform: uppercase;
}

.materia-card {
    background: linear-gradient(145deg, #0d1b3e, #152444);
    border: 1px solid #c9a84c44; border-left: 4px solid #c9a84c;
    border-radius: 10px; padding: 20px; margin-bottom: 16px;
}
.materia-card h3 {
    font-family: 'Playfair Display', serif;
    color: #c9a84c; margin: 0 0 10px; font-size: 1.1rem;
}
.materia-card .stats { display: flex; gap: 20px; font-size: .85rem; color: #a0aec0; }
.materia-card .stat-val { color: #e8c97a; font-weight: 600; font-size: 1.1rem; }

.metric-box {
    background: linear-gradient(145deg, #0d1b3e, #152444);
    border: 1px solid #c9a84c33; border-top: 3px solid #c9a84c;
    border-radius: 10px; padding: 18px 20px; text-align: center;
}
.metric-box .label {
    font-size: .72rem; color: #6b7280;
    text-transform: uppercase; letter-spacing: 1px; margin-bottom: 6px;
}
.metric-box .value {
    font-family: 'Playfair Display', serif;
    font-size: 2rem; color: #c9a84c; font-weight: 700;
}
.metric-box .sub { font-size: .78rem; color: #a0aec0; margin-top: 4px; }

[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #0a1528 0%, #0d1b3e 100%);
    border-right: 1px solid #c9a84c33;
}
[data-testid="stSidebar"] * { color: #e2e8f0 !important; }

.stButton > button {
    background: linear-gradient(135deg, #c9a84c, #e8c97a) !important;
    color: #0d1b3e !important; font-weight: 700 !important;
    border: none !important; border-radius: 6px !important;
}
.stButton > button:hover {
    opacity: .9 !important;
    box-shadow: 0 4px 12px rgba(201,168,76,.35) !important;
}

.gold-divider {
    height: 1px;
    background: linear-gradient(90deg, transparent, #c9a84c, transparent);
    margin: 20px 0;
}
.versiculo {
    background: #0a1528; border-left: 3px solid #c9a84c;
    padding: 12px 18px; border-radius: 0 8px 8px 0;
    font-style: italic; color: #a0aec0; font-size: .85rem; margin: 16px 0;
}
.section-title {
    font-family: 'Playfair Display', serif; color: #c9a84c;
    font-size: 1.3rem; border-bottom: 1px solid #c9a84c44;
    padding-bottom: 8px; margin-bottom: 20px;
}
</style>
""", unsafe_allow_html=True)

# ── Conexión Google Sheets ────────────────────────────────────────────────────

@st.cache_resource
def get_client():
    creds_dict = st.secrets["gcp_service_account"]
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return gspread.authorize(creds)

def get_spreadsheet():
    return get_client().open_by_key(SHEET_ID)

@st.cache_data(ttl=30)
def get_materias():
    return [ws.title for ws in get_spreadsheet().worksheets()]

@st.cache_data(ttl=30)
def get_estudiantes(materia):
    try:
        data = get_spreadsheet().worksheet(materia).get_all_records()
        if not data:
            return pd.DataFrame(columns=["Nombre", "Nota", "Fecha", "Letra"])
        df = pd.DataFrame(data)
        for col in ["Nombre", "Nota", "Fecha", "Letra"]:
            if col not in df.columns:
                df[col] = ""
        df["Nota"] = pd.to_numeric(df["Nota"], errors="coerce").fillna(0)
        return df[["Nombre", "Nota", "Fecha", "Letra"]].dropna(subset=["Nombre"])
    except:
        return pd.DataFrame(columns=["Nombre", "Nota", "Fecha", "Letra"])

def _limpiar_cache():
    get_materias.clear()
    get_estudiantes.clear()

def guardar_estudiantes(materia, df):
    spr = get_spreadsheet()
    try:
        ws = spr.worksheet(materia)
        ws.clear()
    except:
        ws = spr.add_worksheet(title=materia, rows=500, cols=10)
    rows = [["Nombre", "Nota", "Fecha", "Letra"]]
    for _, row in df.iterrows():
        rows.append([str(row["Nombre"]), int(row["Nota"]), str(row["Fecha"]), str(row["Letra"])])
    ws.update(rows, "A1")
    _limpiar_cache()

def crear_materia(nombre):
    if nombre in get_materias():
        return False
    spr = get_spreadsheet()
    ws  = spr.add_worksheet(title=nombre, rows=500, cols=10)
    ws.update([["Nombre", "Nota", "Fecha", "Letra"]], "A1")
    _limpiar_cache()
    return True

def eliminar_materia(nombre):
    if len(get_materias()) <= 1:
        return False
    spr = get_spreadsheet()
    spr.del_worksheet(spr.worksheet(nombre))
    _limpiar_cache()
    return True

def nota_a_letra(nota):
    if nota >= 90: return "A"
    if nota >= 80: return "B"
    if nota >= 70: return "C"
    if nota >= 60: return "D"
    return "F"

# ── Exportar Excel con estilo ─────────────────────────────────────────────────

def exportar_excel():
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    navy_fill  = PatternFill("solid", fgColor="0D1B3E")
    gold_fill  = PatternFill("solid", fgColor="C9A84C")
    white_font = Font(color="FFFFFF", bold=True, name="Calibri", size=11)
    navy_font  = Font(color="0D1B3E", bold=True, name="Calibri", size=11)
    gold_font  = Font(color="C9A84C", bold=True, name="Calibri", size=14)
    border     = Border(
        left=Side(style="thin", color="C9A84C"),
        right=Side(style="thin", color="C9A84C"),
        top=Side(style="thin", color="C9A84C"),
        bottom=Side(style="thin", color="C9A84C"),
    )

    for materia in get_materias():
        df   = get_estudiantes(materia)
        ws   = wb.create_sheet(materia[:31])

        # Título
        ws.merge_cells("A1:D1")
        ws["A1"] = "✝  IGLESIA PENTECOSTAL FUENTE DE GRACIA"
        ws["A1"].fill = navy_fill; ws["A1"].font = gold_font
        ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 30

        ws.merge_cells("A2:D2")
        ws["A2"] = f"REGISTRO ACADÉMICO — {materia.upper()}"
        ws["A2"].fill = gold_fill; ws["A2"].font = navy_font
        ws["A2"].alignment = Alignment(horizontal="center"); ws.row_dimensions[2].height = 22

        ws.merge_cells("A3:D3")
        ws["A3"] = f"Generado: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
        ws["A3"].font = Font(color="888888", italic=True, name="Calibri", size=9)
        ws["A3"].alignment = Alignment(horizontal="center"); ws.row_dimensions[3].height = 16

        for col, h in enumerate(["Nombre", "Nota", "Fecha", "Letra"], 1):
            c = ws.cell(row=4, column=col, value=h)
            c.fill = navy_fill; c.font = white_font
            c.alignment = Alignment(horizontal="center"); c.border = border
        ws.row_dimensions[4].height = 18

        for i, (_, row) in enumerate(df.iterrows()):
            r    = i + 5
            fill = PatternFill("solid", fgColor="F8F6F0" if i%2==0 else "EEE8D5")
            for col, val in enumerate([row["Nombre"], row["Nota"], row["Fecha"], row["Letra"]], 1):
                c = ws.cell(row=r, column=col, value=val)
                c.fill = fill; c.alignment = Alignment(horizontal="center"); c.border = border

        if not df.empty:
            pr = len(df) + 5
            ws.merge_cells(f"A{pr}:C{pr}")
            ws[f"A{pr}"] = "PROMEDIO DEL GRUPO"
            ws[f"A{pr}"].fill = gold_fill; ws[f"A{pr}"].font = navy_font
            ws[f"A{pr}"].alignment = Alignment(horizontal="center")
            ws[f"D{pr}"] = round(df["Nota"].mean(), 1)
            ws[f"D{pr}"].fill = gold_fill; ws[f"D{pr}"].font = navy_font
            ws[f"D{pr}"].alignment = Alignment(horizontal="center")

        ws.column_dimensions["A"].width = 28
        ws.column_dimensions["B"].width = 10
        ws.column_dimensions["C"].width = 16
        ws.column_dimensions["D"].width = 10

    output = io.BytesIO()
    wb.save(output); output.seek(0)
    return output

# ── Gráfica ───────────────────────────────────────────────────────────────────

def grafica(df, materia):
    if df.empty: return None
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(11, 4))
    fig.patch.set_facecolor("#0d1b3e")
    nombres = [str(n)[:12]+"…" if len(str(n))>12 else str(n) for n in df["Nombre"]]
    notas   = df["Nota"].tolist()
    colores = ["#6ee7b7" if n>=90 else "#93c5fd" if n>=80 else
               "#fcd34d" if n>=70 else "#fdba74" if n>=60 else "#fca5a5" for n in notas]
    bars = ax1.bar(nombres, notas, color=colores, edgecolor="#c9a84c", linewidth=.7)
    ax1.set_facecolor("#0a1528"); ax1.tick_params(colors="#a0aec0", labelsize=8)
    ax1.spines[:].set_color("#1e2d4e"); ax1.set_ylim(0, 110)
    ax1.axhline(70, color="#c9a84c", linestyle="--", linewidth=.8, alpha=.6)
    ax1.set_title(f"Notas — {materia[:28]}", color="#c9a84c", fontsize=10, fontweight="bold")
    plt.setp(ax1.xaxis.get_majorticklabels(), rotation=30, ha="right")
    for bar, n in zip(bars, notas):
        ax1.text(bar.get_x()+bar.get_width()/2, bar.get_height()+1,
                 str(int(n)), ha="center", color="white", fontsize=8)
    letras = df["Letra"].value_counts()
    colmap = {"A":"#6ee7b7","B":"#93c5fd","C":"#fcd34d","D":"#fdba74","F":"#fca5a5"}
    pie_cols = [colmap.get(l,"#888") for l in letras.index]
    _, texts, autotexts = ax2.pie(
        letras.values, labels=letras.index, colors=pie_cols,
        autopct="%1.0f%%", startangle=90,
        textprops={"color":"white","fontsize":9},
        wedgeprops={"edgecolor":"#0d1b3e","linewidth":1.5}
    )
    for at in autotexts: at.set_color("#0d1b3e"); at.set_fontweight("bold")
    ax2.set_facecolor("#0a1528")
    ax2.set_title("Distribución de letras", color="#c9a84c", fontsize=10, fontweight="bold")
    plt.tight_layout(pad=2)
    return fig

# ── Header ────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="inst-header">
  <div style="font-size:2.2rem">✝</div>
  <div>
    <h1>Sistema Académico · Instituto Bíblico</h1>
    <p>Iglesia Pentecostal Fuente de Gracia</p>
  </div>
</div>
""", unsafe_allow_html=True)

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### ✝ Navegación")
    st.markdown('<div class="gold-divider"></div>', unsafe_allow_html=True)
    pagina = st.radio("", [
        "📊 Dashboard",
        "📖 Ver Materia",
        "➕ Crear Materia",
        "📈 Estadísticas",
        "📥 Exportar Excel"
    ], label_visibility="collapsed")
    st.markdown('<div class="gold-divider"></div>', unsafe_allow_html=True)
    try:
        materias  = get_materias()
        total_est = sum(len(get_estudiantes(m)) for m in materias)
        st.markdown(f"**{len(materias)}** materias · **{total_est}** estudiantes")
    except:
        st.markdown("Conectando…")
    st.markdown('<div class="gold-divider"></div>', unsafe_allow_html=True)
    st.markdown("""
    <div class="versiculo">
    "Instruye al sabio, y se hará más sabio"<br><strong>— Proverbios 9:9</strong>
    </div>""", unsafe_allow_html=True)

# ── DASHBOARD ─────────────────────────────────────────────────────────────────
if pagina == "📊 Dashboard":
    st.markdown('<div class="section-title">Panel Principal</div>', unsafe_allow_html=True)
    try:
        materias = get_materias()
        iconos   = ["📖","📜","👑","🕊️","🔥","⚓","🌿","🗝️","🛡️","📣"]
        if not materias:
            st.info("No hay materias aún. Ve a **Crear Materia** para empezar.")
        else:
            cols = st.columns(min(3, len(materias)))
            for i, materia in enumerate(materias):
                df       = get_estudiantes(materia)
                total    = len(df)
                promedio = round(df["Nota"].mean(), 1) if total > 0 else 0
                maxnota  = int(df["Nota"].max()) if total > 0 else 0
                with cols[i % min(3, len(materias))]:
                    st.markdown(f"""
                    <div class="materia-card">
                      <h3>{iconos[i%len(iconos)]} {materia}</h3>
                      <div class="stats">
                        <div><div style="font-size:.72rem;color:#6b7280;text-transform:uppercase;letter-spacing:1px">Estudiantes</div><div class="stat-val">{total}</div></div>
                        <div><div style="font-size:.72rem;color:#6b7280;text-transform:uppercase;letter-spacing:1px">Promedio</div><div class="stat-val">{promedio}</div></div>
                        <div><div style="font-size:.72rem;color:#6b7280;text-transform:uppercase;letter-spacing:1px">Más alta</div><div class="stat-val">{maxnota}</div></div>
                      </div>
                    </div>""", unsafe_allow_html=True)
    except Exception as e:
        st.error(f"Error conectando con Google Sheets: {e}")

# ── VER MATERIA ───────────────────────────────────────────────────────────────
elif pagina == "📖 Ver Materia":
    try:
        materias = get_materias()
        if not materias:
            st.warning("No hay materias. Crea una primero.")
        else:
            materia_sel = st.selectbox("Selecciona una materia", materias)
            df = get_estudiantes(materia_sel)
            st.markdown(f'<div class="section-title">📖 {materia_sel}</div>', unsafe_allow_html=True)

            total     = len(df)
            prom      = round(df["Nota"].mean(), 1) if total > 0 else 0
            maxn      = int(df["Nota"].max()) if total > 0 else 0
            aprobados = len(df[df["Nota"] >= 70]) if total > 0 else 0

            c1, c2, c3, c4 = st.columns(4)
            for col, label, val, sub in zip(
                [c1,c2,c3,c4],
                ["Estudiantes","Promedio","Nota más alta","Aprobados"],
                [total, prom, maxn, aprobados],
                ["inscritos","del grupo","registrada",f"de {total}"]
            ):
                col.markdown(f"""
                <div class="metric-box">
                  <div class="label">{label}</div>
                  <div class="value">{val}</div>
                  <div class="sub">{sub}</div>
                </div>""", unsafe_allow_html=True)

            st.markdown('<div class="gold-divider"></div>', unsafe_allow_html=True)

            with st.expander("➕ Agregar / Actualizar Estudiante", expanded=False):
                col1, col2 = st.columns(2)
                with col1:
                    nuevo_nombre = st.text_input("Nombre completo")
                with col2:
                    nueva_nota = st.slider("Nota", 0, 100, 75)
                if st.button("💾 Guardar Estudiante"):
                    if nuevo_nombre.strip():
                        letra  = nota_a_letra(nueva_nota)
                        nueva  = pd.DataFrame([{"Nombre": nuevo_nombre.strip(), "Nota": nueva_nota,
                                                "Fecha": datetime.now().strftime("%d/%m/%Y"), "Letra": letra}])
                        df = df[df["Nombre"].str.lower() != nuevo_nombre.strip().lower()]
                        df = pd.concat([df, nueva], ignore_index=True)
                        with st.spinner("Guardando en Google Sheets…"):
                            guardar_estudiantes(materia_sel, df)
                        st.success(f"✅ {nuevo_nombre} guardado · Letra {letra}")
                        st.rerun()
                    else:
                        st.error("Escribe un nombre.")

            buscar = st.text_input("🔍 Buscar estudiante", placeholder="Escribe un nombre…")
            df_vis = df[df["Nombre"].str.contains(buscar, case=False, na=False)] if buscar else df

            if df_vis.empty:
                st.info("No hay estudiantes en esta materia aún.")
            else:
                st.dataframe(df_vis.style.format({"Nota": "{:.0f}"}),
                             use_container_width=True, hide_index=True)
                st.markdown('<div class="gold-divider"></div>', unsafe_allow_html=True)
                a_eliminar = st.selectbox("Eliminar estudiante", ["— selecciona —"] + df["Nombre"].tolist())
                if a_eliminar != "— selecciona —":
                    if st.button(f"🗑️ Eliminar a {a_eliminar}"):
                        df = df[df["Nombre"] != a_eliminar]
                        with st.spinner("Guardando cambios…"):
                            guardar_estudiantes(materia_sel, df)
                        st.success(f"'{a_eliminar}' eliminado.")
                        st.rerun()

            st.markdown('<div class="gold-divider"></div>', unsafe_allow_html=True)
            if not df.empty:
                fig = grafica(df, materia_sel)
                if fig: st.pyplot(fig, use_container_width=True)

            with st.expander("⚠️ Zona de peligro — Eliminar materia", expanded=False):
                st.warning("Esta acción eliminará la materia y todos sus estudiantes.")
                if st.button("🗑️ Eliminar esta materia"):
                    if eliminar_materia(materia_sel):
                        st.success("Materia eliminada."); st.rerun()
                    else:
                        st.error("No puedes eliminar la única materia.")
    except Exception as e:
        st.error(f"Error: {e}")

# ── CREAR MATERIA ─────────────────────────────────────────────────────────────
elif pagina == "➕ Crear Materia":
    st.markdown('<div class="section-title">➕ Nueva Materia</div>', unsafe_allow_html=True)
    st.markdown("""
    <div class="versiculo">
    Sugerencias: Teología Sistemática · Hermenéutica · Escatología ·
    Liderazgo Cristiano · Evangelismo · Homilética · Ética Cristiana
    </div>""", unsafe_allow_html=True)
    nombre_materia = st.text_input("Nombre de la materia", placeholder="ej: Hermenéutica")
    if st.button("✅ Crear Materia"):
        if nombre_materia.strip():
            with st.spinner("Creando materia en Google Sheets…"):
                if crear_materia(nombre_materia.strip()):
                    st.success(f"✅ Materia '{nombre_materia}' creada."); st.balloons()
                else:
                    st.warning("Ya existe una materia con ese nombre.")
        else:
            st.error("Escribe el nombre de la materia.")
    st.markdown('<div class="gold-divider"></div>', unsafe_allow_html=True)
    st.markdown("**Materias actuales:**")
    try:
        for m in get_materias():
            st.markdown(f"- 📖 {m}")
    except:
        st.info("Cargando…")

# ── ESTADÍSTICAS ──────────────────────────────────────────────────────────────
elif pagina == "📈 Estadísticas":
    st.markdown('<div class="section-title">📈 Estadísticas Generales</div>', unsafe_allow_html=True)
    try:
        resumen = []
        for m in get_materias():
            df = get_estudiantes(m)
            if not df.empty:
                resumen.append({
                    "Materia": m, "Estudiantes": len(df),
                    "Promedio": round(df["Nota"].mean(), 1),
                    "Más alta": int(df["Nota"].max()),
                    "Más baja": int(df["Nota"].min()),
                    "Aprobados": len(df[df["Nota"] >= 70]),
                })
        if not resumen:
            st.info("Aún no hay datos registrados.")
        else:
            df_res = pd.DataFrame(resumen)
            st.dataframe(df_res, use_container_width=True, hide_index=True)
            st.markdown('<div class="gold-divider"></div>', unsafe_allow_html=True)
            fig, ax = plt.subplots(figsize=(10, 4))
            fig.patch.set_facecolor("#0d1b3e"); ax.set_facecolor("#0a1528")
            bars = ax.bar([m[:18] for m in df_res["Materia"]], df_res["Promedio"],
                          color="#c9a84c", edgecolor="#0d1b3e", linewidth=1)
            ax.axhline(70, color="#fca5a5", linestyle="--", linewidth=1, label="Mínimo (70)")
            ax.set_title("Promedio por Materia", color="#c9a84c", fontsize=12, fontweight="bold")
            ax.tick_params(colors="#a0aec0"); ax.spines[:].set_color("#1e2d4e"); ax.set_ylim(0, 110)
            ax.legend(facecolor="#0a1528", labelcolor="#a0aec0", fontsize=8)
            for bar, val in zip(bars, df_res["Promedio"]):
                ax.text(bar.get_x()+bar.get_width()/2, bar.get_height()+1,
                        str(val), ha="center", color="white", fontsize=9, fontweight="bold")
            plt.setp(ax.xaxis.get_majorticklabels(), rotation=20, ha="right", color="#a0aec0")
            plt.tight_layout(); st.pyplot(fig, use_container_width=True)
    except Exception as e:
        st.error(f"Error: {e}")

# ── EXPORTAR ──────────────────────────────────────────────────────────────────
elif pagina == "📥 Exportar Excel":
    st.markdown('<div class="section-title">📥 Exportar Reporte</div>', unsafe_allow_html=True)
    st.markdown("""
    <div class="versiculo">
    El archivo incluye diseño institucional, encabezados en azul marino y dorado,
    y promedio automático por materia.
    </div>""", unsafe_allow_html=True)
    try:
        materias = get_materias()
        for m in materias:
            df = get_estudiantes(m)
            c1, c2, c3 = st.columns([3,1,1])
            c1.markdown(f"📖 **{m}**")
            c2.markdown(f"{len(df)} estudiantes")
            c3.markdown(f"Prom: **{round(df['Nota'].mean(),1) if len(df)>0 else 0}**")
        st.markdown('<div class="gold-divider"></div>', unsafe_allow_html=True)
        if st.button("📊 Generar Excel"):
            with st.spinner("Generando reporte…"):
                excel_data = exportar_excel()
            st.download_button(
                label="⬇️ Descargar Reporte Excel Completo",
                data=excel_data,
                file_name=f"Reporte_FuenteDeGracia_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    except Exception as e:
        st.error(f"Error: {e}")
