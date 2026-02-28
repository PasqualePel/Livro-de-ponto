import streamlit as st
from datetime import datetime, date, timedelta
import pandas as pd
from io import BytesIO
import calendar

# Configurazione della pagina
st.set_page_config(
    page_title="Livro de Ponto",
    page_icon="⛪",
    layout="wide"
)

# Inizializza la sessione per memorizzare i dati
if 'registos' not in st.session_state:
    st.session_state.registos = []

# Feste Nazionali del Mozambico
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

# SIDEBAR - Menu laterale con i mesi
st.sidebar.title("📅 Navegação Mensal")
meses = {
    1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril",
    5: "Maio", 6: "Junho", 7: "Julho", 8: "Agosto",
    9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro"
}

ano_atual = datetime.now().year
mes_selecionado = st.sidebar.selectbox(
    "Selecione o mês:",
    options=list(meses.keys()),
    format_func=lambda x: meses[x],
    index=datetime.now().month - 1
)

st.sidebar.markdown("---")
st.sidebar.markdown(f"**Ano:** {ano_atual}")
st.sidebar.markdown(f"**Mês:** {meses[mes_selecionado]}")

# Calcola quanti giorni ha il mese selezionato
num_dias_mes = calendar.monthrange(ano_atual, mes_selecionado)[1]

st.sidebar.markdown(f"**Dias no mês:** {num_dias_mes}")
st.sidebar.markdown("---")

# Mostra riepilogo ore totali del mese nella sidebar
registos_mes_sidebar = [r for r in st.session_state.registos 
                        if r['Mês'] == mes_selecionado and r['Ano'] == ano_atual]
if registos_mes_sidebar:
    total_minutos_sidebar = 0
    for r in registos_mes_sidebar:
        if "h" in r['Horas']:
            parts = r['Horas'].replace('m', '').split('h')
            total_minutos_sidebar += int(parts[0]) * 60 + int(parts[1].strip())
    
    total_h_sidebar = total_minutos_sidebar // 60
    total_m_sidebar = total_minutos_sidebar % 60
    st.sidebar.success(f"**Total Horas:**\n{total_h_sidebar}h {total_m_sidebar:02d}m")
else:
    st.sidebar.info("Nenhum registo este mês")

# CONTENUTO PRINCIPALE
st.title("⛪ Paróquia SS. Trindade")
st.header("Livro de Ponto")
st.markdown("**Pároco:** Pe. Pasquale Peluso  \n**Secretária:** Yolanda Facitela Clávio")
st.markdown("---")

# Sezione inserimento nuovo registo
st.subheader(f"📝 Novo Registo - {meses[mes_selecionado]} {ano_atual}")

col1, col2, col3 = st.columns(3)

with col1:
    # Selezione del GIORNO del mese selezionato
    dia_selecionado = st.selectbox(
        "Dia do mês",
        options=list(range(1, num_dias_mes + 1)),
        index=min(datetime.now().day - 1, num_dias_mes - 1) if mes_selecionado == datetime.now().month else 0
    )
    
    # Costruisce la data completa
    data_completa = date(ano_atual, mes_selecionado, dia_selecionado)
    
    # Mostra la data formattata
    st.info(f"📅 {data_completa.strftime('%d/%m/%Y')}")
    
    # Nome del giorno della settimana in portoghese
    dias_semana = ['Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta', 'Sábado', 'Domingo']
    dia_semana = dias_semana[data_completa.weekday()]
    st.caption(f"({dia_semana})")
    
    # Controlla se è festa
    giorno_mese = data_completa.strftime("%d-%m")
    feriado_oggi = feriados.get(giorno_mese, "")
    
    if feriado_oggi:
        st.warning(f"🎉 {feriado_oggi}")

with col2:
    # Orario di entrata
    entrada_input = st.text_input(
        "Hora de Entrada",
        placeholder="08:23",
        key=f"entrada_{mes_selecionado}_{dia_selecionado}"
    )

with col3:
    # Orario di uscita
    saida_input = st.text_input(
        "Hora de Saída",
        placeholder="16:30",
        key=f"saida_{mes_selecionado}_{dia_selecionado}"
    )

# Campo note
notas_input = st.text_input(
    "Notas / Observações",
    value=f"Feriado: {feriado_oggi}" if feriado_oggi else "",
    key=f"notas_{mes_selecionado}_{dia_selecionado}"
)

# Funzione per calcolare le ore lavorate
def calcular_horas(entrada_str, saida_str):
    try:
        entrada = datetime.strptime(entrada_str, "%H:%M").time()
        saida = datetime.strptime(saida_str, "%H:%M").time()
        
        min_entrada = entrada.hour * 60 + entrada.minute
        min_saida = saida.hour * 60 + saida.minute
        diff_min = min_saida - min_entrada
        
        if diff_min > 0:
            h = diff_min // 60
            m = diff_min % 60
            return f"{h}h {m:02d}m", True
        else:
            return "Erro: saída antes da entrada", False
    except:
        return "Formato inválido (use HH:MM)", False

# Pulsante per salvare
if st.button("✅ Guardar Registo", type="primary", use_container_width=True):
    if entrada_input and saida_input:
        horas_trabalhadas, valido = calcular_horas(entrada_input, saida_input)
        
        if valido:
            # Controlla se esiste già un registo per questo giorno
            registo_esistente = None
            for idx, r in enumerate(st.session_state.registos):
                if r['Data'] == data_completa.strftime("%d/%m/%Y"):
                    registo_esistente = idx
                    break
            
            novo_registo = {
                "Data": data_completa.strftime("%d/%m/%Y"),
                "Entrada": entrada_input,
                "Saída": saida_input,
                "Horas": horas_trabalhadas,
                "Notas": notas_input,
                "Mês": mes_selecionado,
                "Ano": ano_atual,
                "DiaSemana": dia_semana
            }
            
            if registo_esistente is not None:
                st.session_state.registos[registo_esistente] = novo_registo
                st.success(f"✅ Registo atualizado! ({data_completa.strftime('%d/%m/%Y')}) - Horas: {horas_trabalhadas}")
            else:
                st.session_state.registos.append(novo_registo)
                st.success(f"✅ Registo guardado! ({data_completa.strftime('%d/%m/%Y')}) - Horas: {horas_trabalhadas}")
            st.balloons()
        else:
            st.error(horas_trabalhadas)
    else:
        st.warning("⚠️ Por favor, insira a hora de entrada e saída.")

st.markdown("---")

# Mostra tutti i giorni del mese con i dati
st.subheader(f"📊 Registos de {meses[mes_selecionado]} {ano_atual}")

# Filtra e ordina i dati per il mese selezionato
registos_mes = [r for r in st.session_state.registos 
                if r['Mês'] == mes_selecionado and r['Ano'] == ano_atual]

# Ordina per data
registos_mes_sorted = sorted(registos_mes, key=lambda x: datetime.strptime(x['Data'], "%d/%m/%Y"))

if registos_mes_sorted:
    # Crea DataFrame per visualizzazione
    df = pd.DataFrame(registos_mes_sorted)
    df_display = df[['Data', 'DiaSemana', 'Entrada', 'Saída', 'Horas', 'Notas']]
    
    st.dataframe(df_display, use_container_width=True, hide_index=True)
    
    # Calcola totale ore del mese
    total_minutos = 0
    for r in registos_mes_sorted:
        if "h" in r['Horas']:
            parts = r['Horas'].replace('m', '').split('h')
            total_minutos += int(parts[0]) * 60 + int(parts[1].strip())
    
    total_h = total_minutos // 60
    total_m = total_minutos % 60
    
    col_metric1, col_metric2 = st.columns(2)
    with col_metric1:
        st.metric("Total de Horas Trabalhadas", f"{total_h}h {total_m:02d}m")
    with col_metric2:
        st.metric("Dias Registados", len(registos_mes_sorted))
    
    # PULSANTE PER SCARICARE IL PDF DEL MESE
    st.markdown("---")
    st.subheader("📥 Exportar para Excel")
    
    # Crea Excel in memoria
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_export = df[['Data', 'DiaSemana', 'Entrada', 'Saída', 'Horas', 'Notas']].copy()
        df_export.columns = ['Data', 'Dia da Semana', 'Entrada', 'Saída', 'Horas Trabalhadas', 'Notas']
        
        df_export.to_excel(writer, sheet_name=meses[mes_selecionado], index=False, startrow=7)
        
        # Accede al foglio per aggiungere intestazione
        worksheet = writer.sheets[meses[mes_selecionado]]
        worksheet['A1'] = 'Paróquia SS. Trindade'
        worksheet['A2'] = 'Livro de Ponto'
        worksheet['A3'] = f'Pároco: Pe. Pasquale Peluso'
        worksheet['A4'] = 'Secretária: Yolanda Facitela Clávio'
        worksheet['A5'] = f'Mês: {meses[mes_selecionado]} {ano_atual}'
        worksheet['A6'] = f'Total de Horas: {total_h}h {total_m:02d}m'
        
        # Formattazione
        from openpyxl.styles import Font, Alignment
        worksheet['A1'].font = Font(bold=True, size=14)
        worksheet['A2'].font = Font(bold=True, size=12)
        
    excel_data = output.getvalue()
    
    st.download_button(
        label=f"📄 Baixar Excel - {meses[mes_selecionado]} {ano_atual}",
        data=excel_data,
        file_name=f"Livro_Ponto_{meses[mes_selecionado]}_{ano_atual}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
    
    st.info("💡 Abra o Excel, imprima e entregue para assinatura da Yolanda.")
    
else:
    st.info(f"📝 Nenhum registo para {meses[mes_selecionado]} {ano_atual}.")
    st.markdown(f"Use o formulário acima para adicionar os dias trabalhados (1 a {num_dias_mes}).")

# Footer
st.markdown("---")
st.caption("Desenvolvido para a Paróquia SS. Trindade - Maputo, Moçambique")
