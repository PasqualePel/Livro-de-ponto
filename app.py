import streamlit as st
from datetime import datetime, date
import pandas as pd
import calendar
import requests
import json
from utils import feriados, meses, dias_semana_pt, calc_horas, gerar_pdf, gerar_excel

st.set_page_config(page_title="Livro de Ponto", page_icon="⛪", layout="wide")

SCRIPT_URL = "https://script.google.com/macros/s/AKfycbzK268XUzyLr4TJlksHJf1k8GEPyOG7DrwItH8iw_dSdb6ysVShyEQlsau2lth27kKY/exec"
SHEET_ID = "1KhoFXD3oB91U-gh1zy1-YCK4Kug4XfdlpwttXPU6KAY"
COLS = ["Data","Dia da Semana","Entrada","Saída","Horas Trabalhadas","Notas"]

@st.cache_data(ttl=30)
def carregar_dados():
    try:
        url = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv"
        return pd.read_csv(url)
    except:
        return pd.DataFrame(columns=COLS+["Mês","Ano"])

def guardar_registo(reg):
    try:
        requests.post(SCRIPT_URL, data=json.dumps(reg),
            headers={"Content-Type":"application/json"},
            timeout=15, allow_redirects=True)
        return True
    except Exception as e:
        st.error(f"Erro: {e}")
        return False

def total_min(df):
    t = 0
    for h in df["Horas Trabalhadas"]:
        h = str(h)
        if "h" in h:
            p = h.replace("m","").split("h")
            try: t += int(p[0])*60+int(p[1].strip())
            except: pass
    return t

# SIDEBAR
st.sidebar.title("📅 Meses")
ano = datetime.now().year
mes = st.sidebar.selectbox("Selecione o mês:",
    options=list(meses.keys()),
    format_func=lambda x: meses[x],
    index=datetime.now().month-1)
st.sidebar.markdown(f"### {meses[mes]} {ano}")
st.sidebar.markdown("---")
num_dias = calendar.monthrange(ano, mes)[1]

df0 = carregar_dados()
if not df0.empty and "Mês" in df0.columns:
    dm0 = df0[(df0["Mês"].astype(str)==str(mes))&(df0["Ano"].astype(str)==str(ano))]
    if not dm0.empty:
        t0 = total_min(dm0)
        st.sidebar.success(f"⏱️ **Total:** {t0//60}h {t0%60:02d}m")
        st.sidebar.info(f"📅 **Dias:** {len(dm0)}")
    else:
        st.sidebar.info("Nenhum registo este mês")

# CABEÇALHO
st.title("⛪ Paróquia SS. Trindade")
st.subheader("Livro de Ponto — Yolanda Facitela Clávio")
st.markdown("**Pároco:** Pe. Pasquale Peluso &nbsp;|&nbsp; **Secretária:** Yolanda Facitela Clávio")
st.markdown("---")

# FORMULÁRIO
st.subheader(f"📝 Novo Registo — {meses[mes]} {ano}")
c1, c2, c3 = st.columns(3)

with c1:
    dia = st.selectbox("Dia do mês",
        options=list(range(1, num_dias+1)),
        index=min(datetime.now().day-1, num_dias-1) if mes==datetime.now().month else 0,
        key=f"dia_{mes}")
    data_obj = date(ano, mes, dia)
    dsem = dias_semana_pt[data_obj.weekday()]
    st.info(f"📅 **{data_obj.strftime('%d/%m/%Y')}**\n\n{dsem}")
    fer = feriados.get(data_obj.strftime("%d-%m"), "")
    if fer:
        st.warning(f"🎉 Feriado: {fer}")

with c2:
    entrada = st.text_input("⏰ Hora de Entrada", placeholder="08:00", key=f"ent_{mes}_{dia}")

with c3:
    saida = st.text_input("⏰ Hora de Saída", placeholder="16:30", key=f"sai_{mes}_{dia}")

notas = st.text_input("📝 Notas", value=f"Feriado: {fer}" if fer else "", key=f"not_{mes}_{dia}")

if st.button("✅ Guardar Registo", type="primary", use_container_width=True):
    if entrada and saida:
        horas, ok = calc_horas(entrada, saida)
        if ok:
            reg = {"Data":data_obj.strftime("%d/%m/%Y"),"DiaSemana":dsem,
                   "Entrada":entrada,"Saida":saida,"Horas":horas,
                   "Notas":notas,"Mes":mes,"Ano":ano}
            with st.spinner("A guardar..."):
                if guardar_registo(reg):
                    st.success(f"✅ Guardado! {data_obj.strftime('%d/%m/%Y')} — {horas}")
                    st.balloons()
                    st.cache_data.clear()
        else:
            st.error(horas)
    else:
        st.warning("⚠️ Por favor, insira a hora de entrada e saída.")

st.markdown("---")

# TABELA
st.subheader(f"📊 Registos de {meses[mes]} {ano}")
df_all = carregar_dados()

if not df_all.empty and "Mês" in df_all.columns:
    do_mes = df_all[(df_all["Mês"].astype(str)==str(mes))&(df_all["Ano"].astype(str)==str(ano))].copy()
    if not do_mes.empty:
        do_mes = do_mes.sort_values("Data")
        cols_ok = [c for c in COLS if c in do_mes.columns]
        st.dataframe(do_mes[cols_ok], use_container_width=True, hide_index=True)
        tot = total_min(do_mes)
        m1, m2 = st.columns(2)
        m1.metric("⏱️ Total de Horas", f"{tot//60}h {tot%60:02d}m")
        m2.metric("📅 Dias Trabalhados", len(do_mes))
        st.markdown("---")
        st.subheader("📥 Exportar para Assinar")
        b1, b2 = st.columns(2)
        with b1:
            pdf = gerar_pdf(do_mes, meses[mes], ano, tot//60, tot%60)
            st.download_button(
                label=f"📄 Baixar PDF — {meses[mes]} {ano}",
                data=pdf,
                file_name=f"LivroPonto_{meses[mes]}_{ano}.pdf",
                mime="application/pdf",
                use_container_width=True)
        with b2:
            excel = gerar_excel(do_mes, cols_ok, meses[mes], ano, tot)
            st.download_button(
                label=f"📊 Baixar Excel — {meses[mes]} {ano}",
                data=excel,
                file_name=f"LivroPonto_{meses[mes]}_{ano}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True)
        st.info("💡 Abra o PDF ou Excel, imprima e entregue para assinatura.")
    else:
        st.info(f"📝 Nenhum registo para {meses[mes]} {ano}.")
else:
    st.info("📝 Ainda sem dados. Adicione o primeiro registo acima!")

st.markdown("---")
st.caption("Paróquia SS. Trindade — Maputo, Moçambique")
