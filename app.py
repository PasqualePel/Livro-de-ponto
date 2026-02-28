import streamlit as st
from datetime import datetime, date
import pandas as pd
import calendar
import requests
import json
from utils import (
    feriados, meses, dias_semana_pt,
    calc_horas, gerar_pdf, gerar_excel
)

st.set_page_config(
    page_title="Livro de Ponto - Paróquia SS. Trindade",
    page_icon="⛪",
    layout="wide"
)

SCRIPT_URL = "https://script.google.com/macros/s/AKfycbzK268XUzyLr4TJlksHJf1k8GEPyOG7DrwItH8iw_dSdb6ysVShyEQlsau2lth27kKY/exec"
SHEET_ID = "1KhoFXD3oB91U-gh1zy1-YCK4Kug4XfdlpwttXPU6KAY"
COL_LISTA = ["Data","Dia da Semana","Entrada","Saída","Horas Trabalhadas","Notas"]

@st.cache_data(ttl=30)
def carregar_dados():
    url = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv"
    try:
        df = pd.read_csv(url)
        return df
    except:
        return pd.DataFrame(columns=[
            "Data","Dia da Semana","Entrada",
            "Saída","Horas Trabalhadas","Notas","Mês","Ano"
        ])

def guardar_registo(reg):
    try:
        requests.post(
            SCRIPT_URL,
            data=json.dumps(reg),
            headers={"Content-Type": "application/json"},
            timeout=15,
            allow_redirects=True
        )
        return True
    except Exception as e:
        st.error(f"Erro ao guardar: {e}")
        return False

def calcular_total(df_mes):
    tot = 0
    for h in df_mes["Horas Trabalhadas"]:
        h = str(h)
        if "h" in h:
            p = h.replace("m","").split("h")
            try:
                tot += int(p[0])*60 + int(p[1].strip())
            except:
                pass
    return tot

# ════════════════════════════════════════════════════════════════
# SIDEBAR
# ════════════════════════════════════════════════════════════════
st.sidebar.title("📅 Meses")
ano = datetime.now().year
mes = st.sidebar.selectbox(
    "Selecione o mês:",
    options=list(meses.keys()),
    format_func=lambda x: meses[x],
    index=datetime.now().month - 1
)
st.sidebar.markdown(f"### {meses[mes]} {ano}")
st.sidebar.markdown("---")
num_dias = calendar.monthrange(ano, mes)[1]

df_side = carregar_dados()
if not df_side.empty and "Mês" in df_side.columns:
    side_mes = df_side[
        (df_side["Mês"].astype(str) == str(mes)) &
        (df_side["Ano"].astype(str) == str(ano))
    ]
    if not side_mes.empty:
        tot_s = calcular_total(side_mes)
        st.sidebar.success(f"⏱️ **Total:** {tot_s//60}h {tot_s%60:02d}m")
        st.sidebar.info(f"📅 **Dias:** {len(side_mes)}")
    else:
        st.sidebar.info("Nenhum registo este mês")

# ════════════════════════════════════════════════════════════════
# CABEÇALHO
# ════════════════════════════════════════════════════════════════
st.title("⛪ Paróquia SS. Trindade")
st.subheader("Livro de Ponto — Yolanda Facitela Clávio")
st.markdown(
    "**Pároco:** Pe. Pasquale Peluso &nbsp;|&nbsp; "
    "**Secretária:** Yolanda Facitela Clávio"
)
st.markdown("---")

# ════════════════════════════════════════════════════════════════
# FORMULÁRIO
# ════════════════════════════════════════════════════════════════
st.subheader(f"📝 Novo Registo — {meses[mes]} {ano}")

col1, col2, col3 = st.columns(3)
with col1:
    dia = st.selectbox(
        "Dia do mês",
        options=list(range(1, num_dias+1)),
        index=min(datetime.now().day-1, num_dias-1)
              if mes == datetime.now().month else 0,
        key=f"dia_{mes}"
    )
    data_obj = date(ano, mes, dia)
    dsem = dias_semana_pt[data_obj.weekday()]
    st.info(f"📅 **{data_obj.strftime('%d/%m/%Y')}**\n\n{dsem}")
    fer = feriados.get(data_obj.strftime("%d-%m"), "")
    if fer:
        st.warning(f"🎉 Feriado: {fer}")

with col2:
    entrada = st.text_input(
        "⏰ Hora de Entrada", placeholder="08:00",
        key=f"ent_{mes}_{dia}")

with col3:
    saida = st.text_input(
        "⏰ Hora de Saída", placeholder="16:30",
        key=f"sai_{mes}_{dia}")

notas = st.text_input(
    "📝 Notas",
    value=f"Feriado: {fer}" if fer else "",
    key=f"not_{mes}_{dia}")

if st.button("✅ Guardar Registo", type="primary",
