import streamlit as st
import pandas as pd
from datetime import datetime, time
from streamlit_gsheets import GSheetsConnection
import io

# 1. CONFIGURAZIONE DELLA PAGINA
st.set_page_config(page_title="Livro de Ponto", page_icon="📝", layout="centered")

# 2. INTESTAZIONE VISIBILE NELL'APP
st.title("📝 Livro de Ponto")
st.markdown("### **Paróquia SS. Trindade**")
st.write("**Pároco:** Pe. Pasquale Peluso")
st.write("**Funcionária:** Yolanda Facitela Clávio")
st.divider()

# 3. CONNESSIONE AL DATABASE (Google Sheets)
# Utilizza le credenziali salvate nei "Secrets" di Streamlit
conn = st.connection("gsheets", type=GSheetsConnection)

# Legge i dati esistenti dal foglio Google
try:
    # ttl=0 assicura che i dati siano aggiornati ogni volta che si ricarica
    df_esistente = conn.read(ttl=0)
    # Rimuoviamo eventuali righe completamente vuote
    df_esistente = df_esistente.dropna(how='all')
except Exception:
    # Se il foglio è vuoto o c'è un errore, crea una tabella base
    df_esistente = pd.DataFrame(columns=["Data", "Entrada", "Saída", "Horas", "Obs"])

# 4. ELENCO FESTIVI MOZAMBICO
feriados_mz = {
    "01-01": "Ano Novo",
    "03-02": "Dia dos Heróis Moçambicanos",
    "07-04": "Dia da Mulher Moçambicana",
    "01-05": "Dia do Trabalhador",
    "25-06": "Dia da Independência Nacional",
    "07-09": "Dia da Vitória",
    "25-09": "Dia das Forças Armadas",
    "04-10": "Dia da Paz e Reconciliação",
    "25-12": "Natal / Dia da Família"
}

# 5. AREA DI INSERIMENTO DATI
st.subheader("Registrar Novo Horário")

col1, col2, col3 = st.columns(3)

with col1:
    # Yolanda può scegliere qualsiasi data nel calendario
    data_sel = st.date_input("Escolha a Data", datetime.now())

with col2:
    # Yolanda scrive l'orario di entrata (es. 08:30)
    h_in = st.time_input("Hora de Entrada", time(8, 0))

with col3:
    # Yolanda scrive l'orario di uscita (es. 16:45)
    h_out = st.time_input("Hora de Saída", time(17, 0))

# Controllo se la data scelta è un festivo
dia_mes = data_sel.strftime("%d-%m")
nota_obs = feriados_mz.get(dia_mes, "")
if nota_obs:
    st.info(f"Atenção: {nota_obs} (Feriado)")

# BOTTONE PER SALVARE
if st.button("💾 Salvar no Registro Permanente"):
    # Calcolo della durata del lavoro
    t1 = datetime.combine(data_sel, h_in)
    t2 = datetime.combine(data_sel, h_out)
    diff_secondi = (t2 - t1).total_seconds()
    total_ore = diff_secondi / 3600

    if total_ore <= 0:
        st.error("Erro: A hora de saída deve ser posterior à hora de entrada!")
    else:
        # Prepariamo la nuova riga da aggiungere
        novo_registro = pd.DataFrame([{
            "Data": data_sel.strftime("%d/%m/%Y"),
            "Entrada": h_in.strftime("%H:%M"),
            "Saída": h_out.strftime("%H:%M"),
            "Horas": round(total_ore, 2),
            "Obs": nota_obs
        }])

        # Uniamo i dati vecchi con il nuovo inserimento
        df_finale = pd.concat([df_esistente, novo_registro], ignore_index=True)
        
        # Inviamo i dati aggiornati al Foglio Google
        conn.update(data=df_finale)
        
        st.success(f"Registro do dia {data_sel.strftime('%d/%m/%Y')} salvo com sucesso!")
        # Ricarica l'app per mostrare i dati aggiornati nella tabella sotto
        st.rerun()

# 6. VISUALIZZAZIONE DEI DATI SALVATI
st.divider()
st.subheader("Histórico de Registros")

if not df_esistente.empty:
    # Mostra la tabella ordinata per data (opzionale)
    st.dataframe(df_esistente, use_container_width=True)

    # 7. GENERAZIONE FILE EXCEL PER FIRMA
    def prepare_excel(df):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # Scriviamo i dati partendo dalla riga 5 per lasciare spazio all'intestazione
            df.to_excel(writer, index=False, sheet_name='Livro_de_Ponto', startrow=4)
            
            workbook = writer.book
            sheet = writer.sheets['Livro_de_Ponto']
            
            # Intestazione formale nel file Excel
            sheet['A1'] = "PARÓQUIA SS. TRINDADE - LIVRO DE PONTO"
            sheet['A2'] = f"Funcionária: Yolanda Facitela Clávio"
            sheet['A3'] = f"Pároco: Pe. Pasquale Peluso"
            
            # Linee per le firme in fondo
            last_row = len(df) + 7
            sheet[f'A{last_row}'] = "__________________________"
            sheet[f'A{last_row+1}'] = "Assinatura: Yolanda F. C."
            
            sheet[f'D{last_row}'] = "__________________________"
            sheet[f'D{last_row+1}'] = "Assinatura: Pe. Pasquale"
            
        return output.getvalue()

    excel_file = prepare_excel(df_esistente)
    
    st.download_button(
        label="📥 Baixar Excel para Assinatura",
        data=excel_file,
        file_name=f"Ponto_Mensal_Yolanda_{datetime.now().strftime('%m_%Y')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    
    # Opzione per cancellare tutto (da usare con cautela a fine mese)
    if st.checkbox("Mostrar opção para limpar registros"):
        if st.button("⚠️ APAGAR TODOS OS REGISTROS"):
            df_vuoto = pd.DataFrame(columns=["Data", "Entrada", "Saída", "Horas", "Obs"])
            conn.update(data=df_vuoto)
            st.warning("Todos i registros foram apagados!")
            st.rerun()
else:
    st.info("Ainda não há registros salvos no banco de dados.")
