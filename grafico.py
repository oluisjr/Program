import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import os
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as ExcelImage
from PIL import Image
from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from datetime import datetime
 
# === CONFIGURAÇÕES ===
ARQUIVO_EXCEL = "Treinamentos Pendentes - Junção dos meses.xlsx"
LOGO_PATH = "LogoCSN_Cinza.png"
FAVICON_PATH = "Favicon.png"

MAPA_COLUNAS = {
    "Janeiro": (1, 2), "Fevereiro": (3, 4), "Março": (5, 6), "Abril": (7, 8),
    "Maio": (9, 10), "Junho": (11, 12), "Julho": (13, 14), "Agosto": (15, 16),
    "Setembro": (17, 18), "Outubro": (19, 20), "Novembro": (21, 22), "Dezembro": (23, 24)
}
 
# === FUNÇÕES ===
def carregar_dados(mes1, mes2, area=None):
    col1, col2 = MAPA_COLUNAS[mes1]
    col3, col4 = MAPA_COLUNAS[mes2]
    df = pd.read_excel(ARQUIVO_EXCEL, sheet_name="Base", skiprows=1)
    df = df.iloc[:, [0, col1, col2, col3, col4]]
    df.columns = [
        "Área",
        f"{mes1} (Em dia)",
        f"{mes1} (Vencido)",
        f"{mes2} (Em dia)",
        f"{mes2} (Vencido)"
    ]
    df = df.dropna(subset=["Área"]).reset_index(drop=True)
    df.iloc[:, 1:] = df.iloc[:, 1:].apply(pd.to_numeric, errors="coerce").fillna(0).astype(int)
    if area and area != "Todas":
        df = df[df["Área"] == area].reset_index(drop=True)
    return df
 
def gerar_grafico(df, mes1, mes2, figsize=(16, 0.35)):
    df_plot = df[::-1]
    fig, ax = plt.subplots(figsize=(figsize[0], len(df_plot) * figsize[1]))
    y = np.arange(len(df_plot))
    bar_h = 0.4
    bars1 = ax.barh(y, df_plot[f"{mes1} (Em dia)"], height=bar_h, label=f"{mes1} Em Dia", color="#c5e0b4")
    bars2 = ax.barh(y, df_plot[f"{mes1} (Vencido)"], height=bar_h, left=df_plot[f"{mes1} (Em dia)"], label=f"{mes1} Vencido", color="#ff7357")
    bars3 = ax.barh(y+bar_h, df_plot[f"{mes2} (Em dia)"], height=bar_h, label=f"{mes2} Em Dia", color="#759a64")
    bars4 = ax.barh(y+bar_h, df_plot[f"{mes2} (Vencido)"], height=bar_h, left=df_plot[f"{mes2} (Em dia)"], label=f"{mes2} Vencido", color="#af2d11")
    
    for bars in [bars1, bars2, bars3, bars4]:
        for bar in bars:
            w = bar.get_width()
            if w > 0:
                ax.text(bar.get_x() + w / 2, bar.get_y() + bar.get_height() / 2, str(int(w)),
                        ha='center', va='center', fontsize=7, color='black')
 
    ax.set_yticks(y + bar_h/2)
    ax.set_yticklabels(df_plot["Área"], fontsize=8)
    ax.set_xticks([])
    ax.legend(fontsize=8)
    ax.grid(axis="x", linestyle="--", alpha=0.7)
    plt.tight_layout()
    return fig
 
def salvar_excel(fig):
    fig.savefig("treinamentos_pendentes.png", bbox_inches="tight", dpi=300)
    img = Image.open("treinamentos_pendentes.png")
    img.save("grafico_temp_excel.png")
    wb = load_workbook(ARQUIVO_EXCEL)
    ws = wb["Base"]
    if hasattr(ws, "_images"):
        ws._images.clear()
    ws.add_image(ExcelImage("grafico_temp_excel.png"), "A2")
    wb.save(ARQUIVO_EXCEL)
    os.remove("grafico_temp_excel.png")
    return os.getcwd()
 
def export_pdf(df, fig):
    from datetime import datetime
    pdf_path = f"treinamentos_pendentes_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
    fig.savefig("grafico_pdf_temp.png", bbox_inches="tight", dpi=300)

    c = canvas.Canvas(pdf_path, pagesize=landscape(A4))
    width, height = landscape(A4)

    c.setFont("Helvetica-Bold", 16)
    c.drawCentredString(width/2, height - 30, "Treinamentos Pendentes por Área")
    c.setFont("Helvetica", 8)

    x_start = 40
    y = height - 60
    row_height = 12
    col_widths = [width * 0.4] + [width * 0.15] * (len(df.columns) - 1)

    # Cabeçalho
    x = x_start
    for i, col in enumerate(df.columns):
        c.drawString(x, y, str(col))
        x += col_widths[i]
    y -= row_height

    # Dados
    for _, row in df.iterrows():
        x = x_start
        for i, val in enumerate(row):
            c.drawString(x, y, str(val))
            x += col_widths[i]
        y -= row_height
        if y < 100:
            c.showPage()
            y = height - 50
            c.setFont("Helvetica", 8)

    img_reader = ImageReader("grafico_pdf_temp.png")
    img_w, img_h = img_reader.getSize()
    aspect_ratio = img_w / img_h
    max_graph_width = width - 80
    max_graph_height = y - 40
    graph_width = min(max_graph_width, max_graph_height * aspect_ratio)
    graph_height = graph_width / aspect_ratio
    x_pos = (width - graph_width) / 2
    c.drawImage(img_reader, x_pos, 40, width=graph_width, height=graph_height, preserveAspectRatio=True, mask='auto')
    c.save()
    os.remove("grafico_pdf_temp.png")
    return pdf_path
 
# === INTERFACE ===
# Configurar favicon (48x24 pixels)
favicon = Image.open(FAVICON_PATH).resize((48, 24))
st.set_page_config(
    layout="wide",
    page_title="CSN - Treinamentos",
    page_icon=favicon
)
 
# Estilos CSS fixos (texto sempre branco nos controles)
st.markdown("""
<style>
    /* Botões */
    .stButton>button {
        color: white !important;
    }
    
    /* Radio buttons no sidebar */
    .stRadio>div {
        background-color: #0073f7;
        padding: 10px;
        border-radius: 15px;
    }
    .stRadio>div>label {
        color: white !important;
    }
    
    /* Select boxes */
    .stSelectbox>div>div>select {
        color: white !important;
    }
</style>
""", unsafe_allow_html=True)
 
tema = st.sidebar.radio("Tema", ["Claro", "Escuro"], label_visibility="hidden")
 
bg_color = "#f1f1f1" if tema == "Claro" else "#06061D"
text_color = "#02001A" if tema == "Claro" else "#fffef6"
button_color = "#0073f7" if tema == "Claro" else "#236bbd"

# Estilos CSS dinâmicos
st.markdown(f"""
    <style>
        .main {{ background-color: {bg_color}; }}
        h1, h2, h3, h4, h5, h6, p, .stText {{ color: {text_color} !important; }}
        .stDataFrame {{ color: {text_color}; }}
        .stButton>button {{ background-color: {button_color}; }}
    </style>
""", unsafe_allow_html=True)
 
# Logo e título
st.image(LOGO_PATH, width=180)
st.title("Painel de Treinamentos Pendentes")
 
# Controles
meses = list(MAPA_COLUNAS.keys())
col1, col2, col3 = st.columns(3)
mes1 = col1.selectbox("Escolha o 1º mês", meses, index=4)
mes2 = col2.selectbox("Escolha o 2º mês", meses, index=5)
areas = ["Todas"] + sorted(pd.read_excel(ARQUIVO_EXCEL, sheet_name="Base", skiprows=1).iloc[:,0].dropna().unique().tolist())
area_sel = col3.selectbox("Filtrar por área", areas)
 
# Conteúdo principal
if mes1 == mes2:
    st.warning("Selecione dois meses diferentes.")
else:
    df = carregar_dados(mes1, mes2, area_sel)
    st.dataframe(df, use_container_width=True)
    
    st.markdown("### Gráfico Comparativo")
    fig = gerar_grafico(df, mes1, mes2)
    st.pyplot(fig)
    
    # Botões de exportação
    col1, col2 = st.columns(2)
    with col1:
        if st.button("Salvar no Excel"):
            try:
                caminho = salvar_excel(fig)
                st.success(f"Gráfico salvo no Excel em: {caminho}")
            except Exception as e:
                st.error(f"Erro ao salvar no Excel: {e}")
    
    with col2:
        if st.button("Exportar PDF"):
            try:
                pdf_path = export_pdf(df, fig)
                with open(pdf_path, "rb") as f:
                    st.download_button(
                        "Baixar PDF",
                        f,
                        file_name=f"treinamentos_pendentes_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
                        mime="application/pdf"
                    )
            except Exception as e:
                st.error(f"Erro ao gerar PDF: {e}")
