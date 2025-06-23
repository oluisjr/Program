import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import os
from io import BytesIO
from openpyxl import Workbook
from openpyxl.drawing.image import Image as ExcelImage
from PIL import Image
from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from datetime import datetime
from supabase import create_client, Client

# === SUPABASE ===
url = st.secrets["supabase"]["url"]
key = st.secrets["supabase"]["key"]
supabase: Client = create_client(url, key)

MAPA_COLUNAS = {
    "Janeiro": (1, 2), "Fevereiro": (3, 4), "Março": (5, 6), "Abril": (7, 8),
    "Maio": (9, 10), "Junho": (11, 12), "Julho": (13, 14), "Agosto": (15, 16),
    "Setembro": (17, 18), "Outubro": (19, 20), "Novembro": (21, 22), "Dezembro": (23, 24)
}

# === FUNÇÕES ===
def carregar_dados(mes1, mes2, area=None):
    response = supabase.table("treinamentos").select("*").execute()
    dados = response.data

    df_raw = pd.DataFrame(dados)
    if df_raw.empty:
        st.warning("Nenhum dado encontrado na tabela 'treinamentos'.")
        return pd.DataFrame()

    # Padronizar nomes das colunas e dados
    df_raw.columns = df_raw.columns.str.lower()
    df_raw["mes"] = df_raw["mes"].str.strip().str.upper()
    df_raw["area"] = df_raw["area"].str.strip().str.upper()

    # Filtrar por área se necessário
    if area and area != "Todas":
        area = area.strip().upper()
        df_raw = df_raw[df_raw["area"] == area]

    # Filtrar pelos meses selecionados
    df_raw = df_raw[df_raw["mes"].isin([mes1.upper(), mes2.upper()])]

    if df_raw.empty:
        st.warning("Nenhum dado correspondente aos meses selecionados.")
        return pd.DataFrame()

    # Gerar tabela pivô
    df = df_raw.pivot_table(index='area', columns='mes', values=['em_dia', 'vencido'], fill_value=0)
    df = df.sort_index(axis=1, level=1)
    df.columns = [f"{mes.title()} ({tipo.replace('_', ' ').title()})" for tipo, mes in df.columns]
    df.reset_index(inplace=True)

    if "area" not in df.columns:
        st.error("Erro: coluna 'area' ausente no DataFrame final.")
        st.stop()

    return df

def salvar_registro(area, mes, em_dia, vencido):
    # Padroniza a entrada
    area = area.strip().upper()
    mes = mes.strip().upper()

    # Verifica se já existe um registro com a mesma área e mês
    resultado = supabase.table("treinamentos").select("id").match({"area": area, "mes": mes}).execute()

    if resultado.data:
        # Atualiza o registro existente
        id_existente = resultado.data[0]['id']
        supabase.table("treinamentos").update({
            "em_dia": em_dia,
            "vencido": vencido
        }).eq("id", id_existente).execute()
    else:
        # Insere novo registro
        supabase.table("treinamentos").insert({
            "area": area,
            "mes": mes,
            "em_dia": em_dia,
            "vencido": vencido
        }).execute()

def gerar_grafico(df, mes1, mes2):
    df_plot = df[::-1].copy()
    fig, ax = plt.subplots(figsize=(16, len(df_plot) * 0.35))
    y = np.arange(len(df_plot))
    bar_h = 0.4

    for idx, (mes, cor1, cor2) in enumerate([(mes1, "#c5e0b4", "#ff7357"), (mes2, "#759a64", "#af2d11")]):
        col_em_dia = f"{mes} (Em Dia)"
        col_vencido = f"{mes} (Vencido)"
        offset = idx * bar_h

        if col_em_dia not in df_plot.columns or col_vencido not in df_plot.columns:
            st.warning(f"Colunas para o mês '{mes}' não foram encontradas.")
            continue

        em_dia_vals = df_plot[col_em_dia]
        vencido_vals = df_plot[col_vencido]

        ax.barh(y + offset, em_dia_vals, height=bar_h, label=f"{mes} Em Dia", color=cor1)
        ax.barh(y + offset, vencido_vals, height=bar_h, left=em_dia_vals, label=f"{mes} Vencido", color=cor2)

        for i, (em, ven) in enumerate(zip(em_dia_vals, vencido_vals)):
            if em > 0:
                ax.text(em / 2, i + offset, str(int(em)), ha="center", va="center", fontsize=7)
            if ven > 0:
                ax.text(em + ven / 2, i + offset, str(int(ven)), ha="center", va="center", fontsize=7)

    ax.set_yticks(y + bar_h / 2)
    if "area" in df_plot.columns:
        ax.set_yticklabels(df_plot["area"], fontsize=8)
    else:
        st.error("Coluna 'area' não encontrada em df_plot.")
        st.stop()
    ax.set_xticks([])
    ax.legend(fontsize=8)
    ax.grid(axis="x", linestyle="--", alpha=0.7)
    plt.tight_layout()
    return fig

def exportar_para_excel(df, fig):
    fig.savefig("grafico_temp.png", bbox_inches="tight", dpi=300)
    wb = Workbook()
    ws = wb.active
    ws.title = "Base"
    ws.append(list(df.columns))
    for linha in df.itertuples(index=False):
        ws.append(list(linha))
    img = ExcelImage("grafico_temp.png")
    img.anchor = "A2"
    ws.add_image(img)
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    os.remove("grafico_temp.png")
    return output

def export_pdf(df, fig):
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
    x = x_start
    for i, col in enumerate(df.columns):
        c.drawString(x, y, str(col))
        x += col_widths[i]
    y -= row_height
    for _, row in df.iterrows():
        x = x_start
        for i, val in enumerate(row):
            if isinstance(val, float) and val.is_integer():
                val = int(val)
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
st.set_page_config(layout="wide", page_title="CSN - Treinamentos")
st.title("Painel de Treinamentos Pendentes")

if "autenticado" not in st.session_state:
    st.session_state["autenticado"] = False

meses = list(MAPA_COLUNAS.keys())
col1, col2, col3 = st.columns(3)
mes1 = col1.selectbox("Escolha o 1º mês", meses, index=4)
mes2 = col2.selectbox("Escolha o 2º mês", meses, index=5)
dados_supabase = supabase.table("treinamentos").select("area").execute().data
areas_distintas = sorted(set(d['area'].strip().upper() for d in dados_supabase))
areas = ["Todas"] + areas_distintas
area_sel = col3.selectbox("Filtrar por área", areas)

if mes1 == mes2:
    st.warning("Selecione dois meses diferentes.")
else:
    df = carregar_dados(mes1, mes2, area_sel)
    if not df.empty:
        st.dataframe(df, use_container_width=True)
        st.markdown("### Gráfico Comparativo")
        fig = gerar_grafico(df, mes1, mes2)
        st.pyplot(fig)
        col1, col2 = st.columns(2)
        with col1:
            if st.button("Salvar no Excel"):
                try:
                    excel_data = exportar_para_excel(df, fig)
                    st.download_button("Baixar Excel", excel_data, file_name="treinamentos_atualizado.xlsx")
                except Exception as e:
                    st.error(f"Erro ao salvar no Excel: {e}")
        with col2:
            if st.button("Exportar PDF"):
                try:
                    pdf_path = export_pdf(df, fig)
                    with open(pdf_path, "rb") as f:
                        st.download_button("Baixar PDF", f, file_name=pdf_path, mime="application/pdf")
                except Exception as e:
                    st.error(f"Erro ao gerar PDF: {e}")

# Carregar lista de áreas distintas
dados_supabase = supabase.table("treinamentos").select("area").execute().data

# Garantir que as áreas sejam padronizadas e ordenadas
areas_distintas = sorted(set(d['area'].strip().upper() for d in dados_supabase))
areas = ["Todas"] + areas_distintas

# === ÁREA PROTEGIDA PARA EDIÇÃO ===
with st.expander("Editar dados (restrito)", expanded=st.session_state["autenticado"]):
    if not st.session_state["autenticado"]:
        senha = st.text_input("Senha de edição", type="password")
        if senha == st.secrets["geral"]["senha_edicao"]:
            st.success("Acesso liberado")
            st.session_state["autenticado"] = True
            st.rerun()
        elif senha:
            st.error("Senha incorreta")
    else:
        if st.button("Encerrar sessão de edição"):
            st.session_state["autenticado"] = False
            st.rerun()

        # Interface de edição
        mes_edicao = st.selectbox("Mês para editar", meses)
        area = st.selectbox("Nome da área", areas)
        em_dia = st.number_input("Em dia", min_value=0, step=1)
        vencido = st.number_input("Vencido", min_value=0, step=1)
        if st.button("Salvar dados"):
            salvar_registro(area, mes_edicao, em_dia, vencido)
            st.success("Dados atualizados com sucesso!")
