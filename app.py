import streamlit as st
from datetime import datetime, date

# Configurazione della pagina
st.set_page_config(page_title="Livro de Ponto", page_icon="⛪")

# Intestazione
st.title("Paróquia SS. Trindade")
st.header("Livro de Ponto")
st.markdown("**Pároco:** Pe. Pasquale Peluso  \n**Secretária:** Yolanda Facitela Clávio")
st.markdown("---")

# Feste Nazionali del Mozambico
feriados = {
    "01-01": "Ano Novo",
    "02-03": "Dia dos Heróis Moçambicanos",
    "04-07": "Dia da Mulher Moçambicana",
    "05-01": "Dia do Trabalhador",
    "06-25": "Dia da Independência Nacional",
    "09-07": "Dia da Vitória",
    "09-25": "Dia das Forças Armadas",
    "10-04": "Paz e Reconciliação",
    "12-25": "Dia da Família"
}

# Inserimento Data (controlla in automatico se è festa)
data_selecionada = st.date_input("Data", value=date.today())
mes_dia = data_selecionada.strftime("%m-%d")
feriado_hoje = feriados.get(mes_dia, "")

# Inserimento Orari
col1, col2 = st.columns(2)
with col1:
    entrada = st.time_input("Hora de Entrada (ex: 08:23)", value=None)
with col2:
    saida = st.time_input("Hora de Saída (ex: 16:30)", value=None)

notas = st.text_input("Notas / Feriado", value=feriado_hoje)

# Pulsante per salvare e calcolare le ore
if st.button("Guardar Registo", type="primary"):
    if entrada and saida:
        # Calcolo delle ore fatte
        min_entr = entrada.hour * 60 + entrada.minute
        min_said = saida.hour * 60 + saida.minute
        diff_min = min_said - min_entr
        
        if diff_min > 0:
            h = diff_min // 60
            m = diff_min % 60
            st.success(f"✅ Registo pronto! Horas trabalhadas: {h}h e {m}m")
            st.info("Nota: Nel prossimo passaggio collegheremo l'app al tuo Foglio Google per memorizzare questi dati.")
        else:
            st.error("A hora de saída deve ser maior que a de entrada.")
    else:
        st.warning("Por favor, insere a hora de entrada e saída.")

st.markdown("---")
st.markdown("### 📥 Imprimir Livro")
# Pulsante che scarica direttamente il tuo foglio Google in Excel
st.markdown("[Clique aqui para baixar o Excel Mensal para assinar](https://docs.google.com/spreadsheets/d/1KhoFXD3oB91U-gh1zy1-YCK4Kug4XfdlpwttXPU6KAY/export?format=xlsx)")
