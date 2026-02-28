import streamlit as st
from datetime import datetime, date
import pandas as pd
import calendar
import requests
import json
from io import BytesIO
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from openpyxl.styles import Font, Alignment

st.set_page_config(page_title="Livro de Ponto", page_icon="⛪", layout="wide")

SCRIPT_URL = "https://script.google.com/macros/s/AKfycbzK268XUzyLr4TJlksHJf1k8GEPyOG7DrwItH8iw_dSdb6ysVShyEQlsau2lth27kKY/exec"
SHEET_ID = "1KhoFXD3oB91U-gh1zy1-YCK4Kug4XfdlpwttXPU6KAY"
COLS = ["Data","Dia da Semana","Entrada","Saída","Horas Trabalhadas","Notas"]

feriados = {
    "01-01":"Ano Novo","03-02":"Dia dos Heróis Moçambicanos",
    "07-04":"Dia da Mulher Moçambicana","01-05":"Dia do Trabalhador",
    "25-06":"Dia da Independência Nacional","07-09":"Dia da Vitória",
    "25-09":"Dia das Forças Armadas","04-10":"Paz e Reconciliação",
    "25-12":"Dia da Família"
}
meses = {1:"Janeiro",2:"Fevereiro",3:"Março",4:"Abril",5:"Maio",6:"Junho",
         7:"Julho",8:"Agosto",9:"Setembro",10:"Outubro",11:"Novembro",12:"Dezembro"}
dias_semana_pt = ['Segunda','Terça','Quarta','Quinta','Sexta','Sábado','Domingo']

def calc_horas(ent, sai):
    try:
        e = datetime.strptime(ent.strip(), "%H:%M")
        s = datetime.strptime(sai.strip(), "%H:%M")
        diff = int((s-e).total_seconds()//60)
        if diff <= 0:
            return "Erro: saída antes da entrada", False
        return f"{diff//60}h {diff%60:02d}m", True
    except:
        return "Formato inválido (use HH:MM)", False

def total_min(df):
    t = 0
    for h in df["Horas Trabalhadas"]:
        h = str(h)
        if "h" in h:
            p = h.replace("m","").split("h")
            try: t += int(p[0])*60+int(p[1].strip())
            except: pass
    return t
def gerar_pdf(do_mes, mes_nome, ano, total_h, total_m):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4,
        rightMargin=2*cm, leftMargin=2*cm,
        topMargin=2*cm, bottomMargin=2*cm)
    styles = getSampleStyleSheet()
    el = []
    t1 = ParagraphStyle('t1',parent=styles['Normal'],fontSize=16,
        fontName='Helvetica-Bold',alignment=TA_CENTER,spaceAfter=6)
    t2 = ParagraphStyle('t2',parent=styles['Normal'],fontSize=13,
        fontName='Helvetica-Bold',alignment=TA_CENTER,spaceAfter=4)
    ti = ParagraphStyle('ti',parent=styles['Normal'],fontSize=11,
        fontName='Helvetica',alignment=TA_CENTER,spaceAfter=2)
    tn = ParagraphStyle('tn',parent=styles['Normal'],fontSize=10,
        fontName='Helvetica',alignment=TA_LEFT,spaceAfter=2)
    tf = ParagraphStyle('tf',parent=styles['Normal'],fontSize=10,
        fontName='Helvetica',alignment=TA_CENTER,spaceAfter=2)
    el.append(Paragraph("Paróquia SS. Trindade", t1))
    el.append(Paragraph("Livro de Ponto — Yolanda Facitela Clávio", t2))
    el.append(Spacer(1, 0.3*cm))
    el.append(Paragraph(f"Mês: {mes_nome} {ano}", ti))
    el.append(Paragraph(f"Total de Horas: {total_h}h {total_m:02d}m", ti))
    el.append(Spacer(1, 0.5*cm))
    el.append(Table([['']],colWidths=[17*cm],
        style=TableStyle([('LINEBELOW',(0,0),(-1,-1),1,colors.black)])))
    el.append(Spacer(1, 0.5*cm))
    cab = ["Data","Dia da Semana","Entrada","Saída","Horas","Notas"]
    dati = [cab]
    for _, row in do_mes.iterrows():
        dati.append([str(row.get("Data","")),str(row.get("Dia da Semana","")),
            str(row.get("Entrada","")),str(row.get("Saída","")),
            str(row.get("Horas Trabalhadas","")),str(row.get("Notas",""))])
    tab = Table(dati,colWidths=[2.5*cm,3.2*cm,2.2*cm,2.2*cm,2.2*cm,4.7*cm],repeatRows=1)
    tab.setStyle(TableStyle([
        ('BACKGROUND',(0,0),(-1,0),colors.HexColor("#2c3e50")),
        ('TEXTCOLOR',(0,0),(-1,0),colors.white),
        ('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),
        ('FONTSIZE',(0,0),(-1,-1),9),
        ('ALIGN',(0,0),(-1,-1),'CENTER'),
        ('FONTNAME',(0,1),(-1,-1),'Helvetica'),
        ('ROWBACKGROUNDS',(0,1),(-1,-1),[colors.white,colors.HexColor("#f2f2f2")]),
        ('GRID',(0,0),(-1,-1),0.5,colors.grey),
        ('BOTTOMPADDING',(0,0),(-1,-1),5),
        ('TOPPADDING',(0,0),(-1,-1),5),
    ]))
    el.append(tab)
    el.append(Spacer(1, 2*cm))
    ass = [
        [Paragraph("_______________________________",tf),Paragraph("_______________________________",tf)],
        [Paragraph("<b>Pároco: Pe. Pasquale Peluso</b>",tf),Paragraph("<b>Secretária: Yolanda Facitela Clávio</b>",tf)],
        [Paragraph("Data: _____ / _____ / _________",tf),Paragraph("Data: _____ / _____ / _________",tf)]
    ]
    tab_ass = Table(ass,colWidths=[8.5*cm,8.5*cm])
    tab_ass.setStyle(TableStyle([
        ('ALIGN',(0,0),(-1,-1),'CENTER'),
        ('VALIGN',(0,0),(-1,-1),'MIDDLE'),
        ('TOPPADDING',(0,0),(-1,-1),8),
    ]))
    el.append(Paragraph("Assinaturas:", tn))
    el.append(Spacer(1, 0.4*cm))
    el.append(tab_ass)
    doc.build(el)
    buffer.seek(0)
    return buffer

def gerar_excel(do_mes, cols_ok, mes_nome, ano, tot):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        do_mes[cols_ok].to_excel(writer, sheet_name=mes_nome, index=False, startrow=7)
        ws = writer.sheets[mes_nome]
        ws["A1"] = "Paróquia SS. Trindade"
        ws["A2"] = "Livro de Ponto — Yolanda Facitela Clávio"
        ws["A3"] = "Pároco: Pe. Pasquale Peluso"
        ws["A4"] = "Secretária: Yolanda Facitela Clávio"
        ws["A5"] = f"Mês: {mes_nome} {ano}"
        ws["A6"] = f"Total de Horas: {tot//60}h {tot%60:02d}m"
        ws["A1"].font = Font(bold=True, size=14)
        ws["A2"].font = Font(bold=True, size=12)
        ws["A1"].alignment = Alignment(horizontal="center")
        ws["A2"].alignment = Alignment(horizontal="center")
    return output.getvalue()

@st.cache_data(ttl=30)
def carregar_dados():
    try:
        url = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv"
        df = pd.read_csv(url)
        return df
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

st.title("⛪ Paróquia SS. Trindade")
st.subheader("Livro de Ponto — Yolanda Facitela Clávio")
st.markdown("**Pároco:** Pe. Pasquale Peluso &nbsp;|&nbsp; **Secretária:** Yolanda Facitela Clávio")
st.markdown("---")
st.subheader(f"📝 Novo Registo — {meses[mes]} {ano}")

c1, c2, c3 = st.columns(3)
with c1:
    dia = st.selectbox("Dia do mês", options=list(range(1,num_dias+1)),
        index=min(datetime.now().day-1,num_dias-1) if mes==datetime.now().month else 0,
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
st.subheader(f"📊 Registos de {meses[mes]} {ano}")

df_all = carregar_dados()
tem_dados = False

if not df_all.empty and "Mês" in df_all.columns:
    do_mes = df_all[(df_all["Mês"].astype(str)==str(mes))&(df_all["Ano"].astype(str)==str(ano))].copy()
    if not do_mes.empty:
        tem_dados = True
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

if not tem_dados:
    st.info(f"📝 Nenhum registo para {meses[mes]} {ano}. Adicione o primeiro registo acima!")

st.markdown("---")
st.caption("Paróquia SS. Trindade — Maputo, Moçambique")
