import streamlit as st
from datetime import datetime, date
import pandas as pd
from io import BytesIO
import calendar
import requests

st.set_page_config(
    page_title="Livro de Ponto - Paróquia SS. Trindade",
    page_icon="⛪",
    layout="wide"
)

# ── Link dello script Google ─────────────────────────────────────
SCRIPT_URL = "https://script.google.com/macros/s/AKfycbzK268XUzyLr4TJlksHJf1k8GEPyOG7DrwItH8iw_dSdb6ysVShyEQlsau2lth27kKY/exec"
SHEET_ID = "1KhoFXD3oB91U-gh1zy1-YCK4Kug4XfdlpwttXPU6KAY"

# ── Feriados Moçambique ──────────────────────────────────────────
feriados = {
    "01-01": "Ano Novo",
    "03-02": "Dia dos Heróis Moçambicanos",
    "07-04": "Dia da Mulher Moçambicana",
    "01-05": "Dia do Trabalhador",
    "25-06": "Dia da Independência Nacional",
    "07-09": "Dia da Vitória",
    "25-09": "Dia das Forças Armadas",
    "04-10": "Paz e Reconciliação",
    "25-12": "Dia da Família"
}

meses = {
    1:"Janeiro", 2:"Fevereiro", 3:"Março", 4:"Abril",
    5:"Maio", 6:"Junho", 7:"Julho", 8:"Agosto",
    9:"Setembro", 10:"Outubro", 11:"Novembro", 12:"Dezembro"
}

dias_semana_pt = ['Segunda','Terça','Quarta','Quinta','Sexta','Sábado','Domingo']

# ── Legge i dati dal foglio Google ──────────────────────────────
@st.cache_data(ttl=30)
def carregar_dados():
    url = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv"
    try:
        df = pd.read_csv(url)
        return df
    except:
        return pd.DataFrame(columns=["Data","Dia da Semana","Entrada","Saída","Horas Trabalhadas","Notas","Mês","Ano"])

# ── Salva i dati tramite Google Apps Script ──────────────────────
def guardar_registo(reg):
    try:
        headers = {"Content-Type": "application/json"}
        resp = requests.post(
            SCRIPT_URL,
            data=str(reg).replace("'", '"'),
            headers=headers,
            timeout=15,
            allow_redirects=True
        )
        return True
    except Exception as e:
        st.error(f"Erro ao guardar: {e}")
        return False

# ── Calcola le ore lavorate ──────────────────────────────────────
def calc_horas(ent, sai):
    try:
        e = datetime.strptime(ent.strip(), "%H:%M")
        s = datetime.strptime(sai.strip(), "%H:%M")
        diff = int((s - e).total_seconds() // 60)
        if diff <= 0:
            return "Erro: saída antes da entrada", False
        return f"{diff//60}h {diff%60:02d}m", True
    except:
        return "Formato inválido (use HH:MM)", False

# ── SIDEBAR ──────────────────────────────────────────────────────
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

# Riepilogo ore nel mese nella sidebar
df_side = carregar_dados()
if not df_side.empty and "Mês" in df_side.columns:
    do_mes_side = df_side[
        (df_side["Mês"].astype(str) == str(mes)) &
        (df_side["Ano"].astype(str) == str(ano))
    ]
    if not do_mes_side.empty:
        tot = 0
        for h in do_mes_side["Horas Trabalhadas"]:
            h = str(h)
            if "h" in h:
                p = h.replace("m","").split("h")
                try: tot += int(p[0])*60 + int(p[1].strip())
                except: pass
        st.sidebar.success(f"⏱️ **Total:** {tot//60}h {tot%60:02d}m")
        st.sidebar.info(f"📅 **Dias registados:** {len(do_mes_side)}")
    else:
        st.sidebar.info("Nenhum registo este mês")

# ── CABEÇALHO PRINCIPAL ──────────────────────────────────────────
st.title("⛪ Paróquia SS. Trindade")
st.subheader("Livro de Ponto")
st.markdown(
    "**Pároco:** Pe. Pasquale Peluso &nbsp;|&nbsp; "
    "**Secretária:** Yolanda Facitela Clávio"
)
st.markdown("---")

# ── FORMULÁRIO DE REGISTO ────────────────────────────────────────
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
        "⏰ Hora de Entrada",
        placeholder="08:00",
        key=f"ent_{mes}_{dia}"
    )

with col3:
    saida = st.text_input(
        "⏰ Hora de Saída",
        placeholder="16:30",
        key=f"sai_{mes}_{dia}"
    )

notas = st.text_input(
    "📝 Notas",
    value=f"Feriado: {fer}" if fer else "",
    key=f"not_{mes}_{dia}"
)

# ── PULSANTE SALVA ───────────────────────────────────────────────
if st.button("✅ Guardar Registo", type="primary", use_container_width=True):
    if entrada and saida:
        horas, ok = calc_horas(entrada, saida)
        if ok:
            reg = {
                "Data": data_obj.strftime("%d/%m/%Y"),
                "DiaSemana": dsem,
                "Entrada": entrada,
                "Saida": saida,
                "Horas": horas,
                "Notas": notas,
                "Mes": mes,
                "Ano": ano
            }
            with st.spinner("A guardar no Google Sheets..."):
                if guardar_registo(reg):
                    st.success(f"✅ Guardado! {data_obj.strftime('%d/%m/%Y')} — {horas}")
                    st.balloons()
                    st.cache_data.clear()
        else:
            st.error(horas)
    else:
        st.warning("⚠️ Por favor, insira a hora de entrada e saída.")

st.markdown("---")

# ── TABELA DO MÊS ────────────────────────────────────────────────
st.subheader(f"📊 Registos de {meses[mes]} {ano}")

df_all = carregar_dados()

if not df_all.empty and "Mês" in df_all.columns:
    do_mes = df_all[
        (df_all["Mês"].astype(str) == str(mes)) &
        (df_all["Ano"].astype(str) == str(ano))
    ].copy()

    if not do_mes.empty:
        do_mes = do_mes.sort_values("Data")
        cols = ["Data","Dia da Semana","Entrada","Saída","Horas Trabalhadas","Notas"]
        cols_ok = [c for c in cols if c in do_mes.columns]
        st.dataframe(do_mes[cols_ok], use_container_width=True, hide_index=True)

        # Totale ore del mese
        tot = 0
        for h in do_mes["Horas Trabalhadas"]:
            h = str(h)
            if "h" in h:
                p = h.replace("m","").split("h")
                try: tot += int(p[0])*60 + int(p[1].strip())
                except: pass

        c1, c2 = st.columns(2)
        c1.metric("⏱️ Total de Horas", f"{tot//60}h {tot%60:02d}m")
        c2.metric("📅 Dias Trabalhados", len(do_mes))

        # ── DOWNLOAD EXCEL ────────────────────────────────────────
        st.markdown("---")
        st.subheader("📥 Exportar Excel para Assinar")

        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            do_mes[cols_ok].to_excel(
                writer, sheet_name=meses[mes], index=False, startrow=7
            )
            ws = writer.sheets[meses[mes]]
            ws["A1"] = "Paróquia SS. Trindade"
            ws["A2"] = "Livro de Ponto"
            ws["A3"] = "Pároco: Pe. Pasquale Peluso"
            ws["A4"] = "Secretária: Yolanda Facitela Clávio"
            ws["A5"] = f"Mês: {meses[mes]} {ano}"
            ws["A6"] = f"Total de Horas: {tot//60}h {tot%60:02d}m"
            from openpyxl.styles import Font, Alignment
            ws["A1"].font = Font(bold=True, size=14)
            ws["A2"].font = Font(bold=True, size=12)
            ws["A1"].alignment = Alignment(horizontal="center")
            ws["A2"].alignment = Alignment(horizontal="center")

        st.download_button(
            label=f"📄 Baixar Excel — {meses[mes]} {ano}",
            data=output.getvalue(),
            file_name=f"LivroPonto_{meses[mes]}_{ano}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        st.info("💡 Abra, imprima e entregue para assinatura da Yolanda.")

    else:
        st.info(f"📝 Nenhum registo para {meses[mes]} {ano}.")
else:
    st.info("📝 Ainda sem dados. Adicione o primeiro registo acima!")

# ── RODAPÉ ───────────────────────────────────────────────────────
st.markdown("---")
st.caption("Paróquia SS. Trindade — Maputo, Moçambique")
