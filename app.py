import streamlit as st
from datetime import datetime, date, time
import pandas as pd
from io import BytesIO

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

# CONTENUTO PRINCIPALE
st.title("⛪ Paróquia SS. Trindade")
st.header("Livro de Ponto")
st.markdown("**Pároco:** Pe. Pasquale Peluso  \n**Secretária:** Yolanda Facitela Clávio")
st.markdown("---")

# Sezione inserimento nuovo registo
st.subheader("📝 Novo Registo")

col1, col2, col3 = st.columns(3)

with col1:
    # Data con formato giorno/mese/anno
    data_input = st.date_input(
        "Data",
        value=date.today(),
        format="DD/MM/YYYY"
    )
    
    # Controlla se è festa
    giorno_mese = data_input.strftime("%d-%m")
    feriado_oggi = feriados.get(giorno_mese, "")
    
    if feriado_oggi:
        st.info(f"🎉 {feriado_hoje}")

with col2:
    # Orario di entrata - tu lo scrivi manualmente
    entrada_input = st.text_input(
        "Hora de Entrada (ex: 08:23)",
        placeholder="08:23"
    )

with col3:
    # Orario di uscita - tu lo scrivi manualmente
    saida_input = st.text_input(
        "Hora de Saída (ex: 16:30)",
        placeholder="16:30"
    )

# Campo note
notas_input = st.text_input(
    "Notas / Observações",
    value=f"Feriado: {feriado_hoje}" if feriado_oggi else ""
)

# Funzione per calcolare le ore lavorate
def calcular_horas(entrada_str, saida_str):
    try:
        # Converte le stringhe in oggetti time
        entrada = datetime.strptime(entrada_str, "%H:%M").time()
        saida = datetime.strptime(saida_str, "%H:%M").time()
        
        # Calcola i minuti totali
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
        return "Formato inválido", False

# Pulsante per salvare
if st.button("✅ Guardar Registo", type="primary", use_container_width=True):
    if entrada_input and saida_input:
        horas_trabalhadas, valido = calcular_horas(entrada_input, saida_input)
        
        if valido:
            # Salva il registo
            novo_registo = {
                "Data": data_input.strftime("%d/%m/%Y"),
                "Entrada": entrada_input,
                "Saída": saida_input,
                "Horas": horas_trabalhadas,
                "Notas": notas_input,
                "Mês": data_input.month,
                "Ano": data_input.year
            }
            st.session_state.registos.append(novo_registo)
            st.success(f"✅ Registo guardado! Horas trabalhadas: {horas_trabalhadas}")
            st.balloons()
        else:
            st.error(horas_trabalhadas)
    else:
        st.warning("⚠️ Por favor, insira a hora de entrada e saída.")

st.markdown("---")

# Mostra i dati del mese selezionato
st.subheader(f"📊 Registos de {meses[mes_selecionado]} {ano_atual}")

# Filtra i dati per il mese selezionato
registos_mes = [r for r in st.session_state.registos 
                if r['Mês'] == mes_selecionado and r['Ano'] == ano_atual]

if registos_mes:
    # Crea DataFrame per visualizzazione
    df = pd.DataFrame(registos_mes)
    df_display = df[['Data', 'Entrada', 'Saída', 'Horas', 'Notas']]
    
    st.dataframe(df_display, use_container_width=True, hide_index=True)
    
    # Calcola totale ore del mese
    total_minutos = 0
    for r in registos_mes:
        if "h" in r['Horas']:
            parts = r['Horas'].replace('m', '').split('h')
            total_minutos += int(parts[0]) * 60 + int(parts[1].strip())
    
    total_h = total_minutos // 60
    total_m = total_minutos % 60
    
    st.metric("Total de Horas Trabalhadas no Mês", f"{total_h}h {total_m:02d}m")
    
    # PULSANTE PER SCARICARE IL PDF DEL MESE
    st.markdown("---")
    st.subheader("📥 Exportar para Excel/PDF")
    
    # Crea Excel in memoria
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Aggiunge intestazione personalizzata
        header_df = pd.DataFrame({
            'A': ['Paróquia SS. Trindade', 
                  'Livro de Ponto',
                  f'Pároco: Pe. Pasquale Peluso',
                  'Secretária: Yolanda Facitela Clávio',
                  f'Mês: {meses[mes_selecionado]} {ano_atual}',
                  '',
                  '']
        })
        
        df_export = df[['Data', 'Entrada', 'Saída', 'Horas', 'Notas']]
        final_df = pd.concat([header_df, df_export], ignore_index=True)
        final_df.to_excel(writer, sheet_name=meses[mes_selecionado], index=False, header=False)
    
    excel_data = output.getvalue()
    
    st.download_button(
        label=f"📄 Baixar Excel - {meses[mes_selecionado]} {ano_atual}",
        data=excel_data,
        file_name=f"Livro_Ponto_{meses[mes_selecionado]}_{ano_atual}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
    
    st.info("💡 Pode abrir o ficheiro Excel, imprimir e entregar para assinatura.")
    
else:
    st.info(f"Nenhum registo para {meses[mes_selecionado]} {ano_atual}. Adicione o primeiro registo acima!")

# Footer
st.markdown("---")
st.caption("Desenvolvido para a Paróquia SS. Trindade - Maputo, Moçambique")
