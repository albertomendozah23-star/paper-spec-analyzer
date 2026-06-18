import streamlit as st
import tempfile, os
from processor import process_pdf

st.set_page_config(
    page_title="Paper Spec Analyzer",
    page_icon="📄",
    layout="centered"
)

# ── Estilos ──────────────────────────────────────────────────────────────────
st.markdown("""
<style>
  .block-container { padding-top: 2rem; max-width: 700px; }
  .title-box {
      background: linear-gradient(135deg, #1a1a2e, #0f3460);
      border-radius: 16px; padding: 28px 32px; margin-bottom: 24px; text-align: center;
  }
  .title-box h1 { color: #FFD700; font-size: 2rem; margin: 0; }
  .title-box p  { color: #ccc; margin: 6px 0 0; font-size: 0.95rem; }
  .info-box {
      background: #F0FFF4; border: 1.5px solid #9AE6B4;
      border-radius: 10px; padding: 12px 16px; margin-top: 8px;
  }
  .feat-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; margin-top: 20px; }
  .feat { background: #F8F9FA; border-radius: 10px; padding: 12px 14px; }
</style>
""", unsafe_allow_html=True)

# ── Header ───────────────────────────────────────────────────────────────────
st.markdown("""
<div class="title-box">
  <h1>📄 Paper Spec Analyzer</h1>
  <p>Sube tu PDF de especificaciones y obtén un Excel con datos y gráficas de capacidad de proceso</p>
</div>
""", unsafe_allow_html=True)

# ── Features ─────────────────────────────────────────────────────────────────
col1, col2, col3, col4 = st.columns(4)
col1.metric("📊", "Datos", "Extraídos")
col2.metric("📈", "Pp / Ppk", "Largo plazo")
col3.metric("📉", "Cp / Cpk", "Corto plazo")
col4.metric("🟢", "Semáforo", "Por índice")

st.divider()

# ── Upload ───────────────────────────────────────────────────────────────────
st.subheader("1️⃣ Sube tu PDF")
uploaded = st.file_uploader(
    "Arrastra o selecciona el PDF de especificaciones",
    type=["pdf"],
    help="Solo archivos .pdf con texto seleccionable"
)

if uploaded:
    st.markdown(f"""
    <div class="info-box">
      📋 <strong>{uploaded.name}</strong> &nbsp;·&nbsp; {uploaded.size/1024:.1f} KB
    </div>
    """, unsafe_allow_html=True)

    st.subheader("2️⃣ Procesar")
    if st.button("⚡ Generar Excel con capacidad de proceso", type="primary", use_container_width=True):
        with st.spinner("Extrayendo datos del PDF..."):
            with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp_pdf:
                tmp_pdf.write(uploaded.read())
                tmp_pdf_path = tmp_pdf.name

        with st.spinner("Calculando Pp, Ppk, Cp, Cpk..."):
            with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp_xlsx:
                tmp_xlsx_path = tmp_xlsx.name

        try:
            with st.spinner("Generando gráficas y Excel..."):
                _, n_rows = process_pdf(tmp_pdf_path, tmp_xlsx_path)

            st.success(f"✅ ¡Listo! Se extrajeron **{n_rows} variables** del PDF.")

            with open(tmp_xlsx_path, "rb") as f:
                excel_bytes = f.read()

            st.subheader("3️⃣ Descargar")
            st.download_button(
                label="📥 Descargar Excel",
                data=excel_bytes,
                file_name=f"{uploaded.name.replace('.pdf','')}_reporte.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                type="primary"
            )
        except Exception as e:
            st.error(f"❌ Error al procesar el PDF: {e}")
        finally:
            try:
                os.remove(tmp_pdf_path)
                os.remove(tmp_xlsx_path)
            except:
                pass

else:
    st.info("👆 Sube un PDF para comenzar")

# ── Footer ───────────────────────────────────────────────────────────────────
st.divider()
st.caption("Paper Spec Analyzer · Desarrollado con Python + Streamlit")
