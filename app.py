import streamlit as st
from datetime import datetime, date
import pandas as pd
from io import BytesIO
import calendar
import requests
import json
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT

st.set_page_config(
    page_title="Livro de Ponto - Paróquia SS. Trindade",
    page_icon="⛪",
    layout="wide"
)

# ── Configuração ─────────────────────────────────────────────────
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
        return pd.DataFrame(columns=[
            "Data","Dia da Semana","Entrada",
            "Saída","Horas Trabalhadas","Notas","Mês","Ano"
        ])

# ── Salva i dati tramite Google Apps Script ──────────────────────
def guardar_registo(reg):
    try:
        headers = {"Content-Type": "application/json"}
        requests.post(
            SCRIPT_URL,
            data=json.dumps(reg),
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

# ── Genera PDF del mese ──────────────────────────────────────────
def gerar_pdf(do_mes, mes_nome, ano, total_h, total_m):
    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=2*cm,
        leftMargin=2*cm,
        topMargin=2*cm,
        bottomMargin=2*cm
    )

    styles = getSampleStyleSheet()
    elementos = []

    estilo_titulo = ParagraphStyle(
        'titulo',
        parent=styles['Normal'],
        fontSize=16,
        fontName='Helvetica-Bold',
        alignment=TA_CENTER,
        spaceAfter=4
    )
    estilo_subtitulo = ParagraphStyle(
        'subtitulo',
        parent=styles['Normal'],
        fontSize=13,
        fontName='Helvetica-Bold',
        alignment=TA_CENTER,
        spaceAfter=4
    )
    estilo_info = ParagraphStyle(
        'info',
        parent=styles['Normal'],
        fontSize=11,
        fontName='Helvetica',
        alignment=TA_CENTER,
        spaceAfter=2
    )
    estilo_normal = ParagraphStyle(
        'normal',
        parent=styles['Normal'],
        fontSize=10,
        fontName='Helvetica',
        alignment=TA_LEFT,
        spaceAfter=2
    )
    estilo_firma = ParagraphStyle(
        'firma',
        parent=styles['Normal'],
        fontSize=10,
        fontName='Helvetica',
        alignment=TA_CENTER,
        spaceAfter=2
    )

    # ── Intestazione PDF ──
    elementos.append(Paragraph("Paróquia SS. Trindade", estilo_titulo))
    elementos.append(Paragraph("Livro de Ponto — Yolanda Facitela Clávio", estilo_subtitulo))
    elementos.append(Spacer(1, 0.2*cm))
    elementos.append(Paragraph(f"Mês: {mes_nome} {ano}", estilo_info))
    elementos.append(Paragraph(
        f"Total de Horas Trabalhadas: {total_h}h {total_m:02d}m", estilo_info
    ))
    elementos.append(Spacer(1, 0.4*cm))

    # Linea separatrice
    elementos.append(Table(
        [['']],
        colWidths=[17*cm],
        style=TableStyle([('LINEBELOW', (0,0), (-1,-1), 1, colors.black)])
    ))
    elementos.append(Spacer(1, 0.4*cm))

    # ── Tabella dati ──
    intestazione = [
        "Data", "Dia da Semana", "Entrada", "Saída", "Horas", "Notas"
    ]
    dati = [intestazione]

    for _, row in do_mes.iterrows():
        dati.append([
            str(row.get("Data", "")),
            str(row.get("Dia da Semana", "")),
            str(row.get("Entrada", "")),
            str(row.get("Saída", "")),
            str(row.get("Horas Trabalhadas", "")),
            str(row.get("Notas", ""))
        ])

    tabela = Table(
        dati,
        colWidths=[2.5*cm, 3.2*cm, 2.2*cm, 2.2*cm, 2.2*cm, 4.7*cm],
        repeatRows=1
    )
    tabela.setStyle(TableStyle([
        ('BACKGROUND',     (0,0), (-1,0),  colors.HexColor("#2c3e50")),
        ('TEXTCOLOR',      (0,0), (-1,0),  colors.white),
        ('FONTNAME',       (0,0), (-1,0),  'Helvetica-Bold'),
        ('FONTSIZE',       (0,0), (-1,0),  9),
        ('ALIGN',          (0,0), (-1,0),  'CENTER'),
        ('FONTNAME',       (0,1), (-1,-1), 'Helvetica'),
        ('FONTSIZE',       (0,1), (-1,-1), 9),
        ('ALIGN',          (0,1), (-1,-1), 'CENTER'),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor("#f2f2f2")]),
        ('GRID',           (0,0), (-1,-1), 0.5, colors.grey),
        ('BOTTOMPADDING',  (0,0), (-1,-1), 5),
        ('TOPPADDING',     (0,0), (-1,-1), 5),
    ]))

    elementos.append(tabela)
    elementos.append(Spacer(1, 1.5*cm))

    # ── Spazio firme ──
    dados_assinatura = [
        [
            Paragraph("_______________________________", estilo_firma),
            Paragraph("_______________________________", estilo_firma)
        ],
        [
            Paragraph("<b>Pároco: Pe. Pasquale Peluso</b>", estilo_firma),
            Paragraph("<b>Secretária: Yolanda Facitela Clávio</b>", estilo_firma)
        ],
        [
            Paragraph("Data: _____ / _____ / _________", estilo_firma),
            Paragraph("Data: _____ / _____ / _________", estilo_firma)
        ]
    ]

    tabela_assinatura = Table(
        dados_assinatura,
        colWidths=[8.5*cm, 8.5*cm]
    )
    tabela_assinatura.setStyle(TableStyle([
        ('ALIGN',        (0,0), (-1,-1), 'CENTER'),
        ('VALIGN',       (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING',   (0,0), (-1,-1), 6),
    ]))

    elementos.append(Paragraph("Assinaturas:", estilo_normal))
    elementos.append(Spacer(1, 0.3*cm))
    elementos.append(tabela_assinatura)

    doc.build(elementos)
    buffer.seek(0)
    return buffer

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
    do_mes_side = df_side[
        (df_side["Mês"].astype(str) == str(mes)) &
        (df_side["Ano"].astype(str) == str(ano))
    ]
    if not do_mes_side.empty:
        tot_s = 0
        for h in do_mes_side["Horas Trabalhadas"]:
            h = str(h)
            if "h" in h:
                p = h.replace("m","").split("h")
                try: tot_s += int(p[0])*60 + int(p[1].strip())
                except: pass
        st.sidebar.success(f"⏱️ **Total:** {tot_s//60}h {tot_s%60:02d}m")
        st.sidebar.info(f"📅 **Dias:** {len(do_mes_side)}")
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
