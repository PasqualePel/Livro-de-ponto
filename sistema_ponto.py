import streamlit as st
import pandas as pd
from datetime import datetime, time
from streamlit_gsheets import GSheetsConnection
import io

# 1. CONFIGURAZIONE DELLA PAGINA
st.set_page_config(page_title="Livro de Ponto - SS. Trindade", page_icon="📝", layout="centered")

# 2. INTESTAZIONE FORMALE
st.title("📝 Livro de Ponto")
st.markdown("### **Paróquia SS. Trindade**")
st.write("**Pároco:** Pe. Pasquale Peluso")
st.write("**Funcionária:** Yolanda Facitela Clávio")
st.divider()

# 3. CONNESSIONE AL DATABASE (Google Sheets)
conn = st.connection("gsheets", type=GSheetsConnection)

def carregar_dados():
    try:
        # Carica i dati dal foglio "Ponto"
        return conn.read(worksheet="Ponto", ttl=0)
    except:
        # Se il foglio è vuoto, crea un DataFrame con le colonne corrette
        return pd.DataFrame(columns=["Data", "Entrada", "Saída"])

# 4. INTERFACCIA DI INSERIMENTO
st.subheader("Registrar Horário")

# Creazione di tre colonne per un inserimento rapido
col1, col2, col3 = st.columns(3)

with col1:
    data_hoje = st.date_input("Data", datetime.now())

with col2:
    # Impostiamo le 08:00 come default per l'entrata
    ora_entrada = st.time_input("Entrada", time(8, 0))

with col3:
    # Impostiamo le 16:30 come default per l'uscita
    ora_saida = st.time_input("Saída", time(16, 30))

if st.button("💾 Salvar Permanentemente"):
    # Prepariamo il nuovo record
    novo_registro = pd.DataFrame([{
        "Data": data_hoje.strftime("%d/%m/%Y"),
        "Entrada": ora_entrada.strftime("%H:%M"),
        "Saída": ora_saida.strftime("%H:%M")
    }])
    
    # Leggiamo i dati attuali e aggiungiamo il nuovo
    df_attuale = carregar_dados()
    df_aggiornato = pd.concat([df_attuale, novo_registro], ignore_index=True)
    
    # Sovrascriviamo il foglio Google
    conn.update(worksheet="Ponto", data=df_aggiornato)
    st.success("✅ Registro salvo com sucesso no Google Sheets!")
    st.balloons()

st.divider()

# 5. STORICO MENSILE
st.subheader("📋 Histórico Mensal")
df_visualizza = carregar_dados()

if not df_visualizza.empty:
    # Mostriamo la tabella con gli ultimi inserimenti in alto
    st.dataframe(df_visualizza.iloc[::-1], use_container_width=True)
else:
    st.info("Nenhum registro encontrado.")
