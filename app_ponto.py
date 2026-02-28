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
        # Legge i dati in tempo reale dal foglio Google
        df = conn.read(ttl=0)
        return df.dropna(how='all')
    except:
        # Se il foglio è nuovo o vuoto, crea la struttura
        return pd.DataFrame(columns=["Data", "Entrada", "Saída", "Horas", "Obs"])

# Carichiamo la lista completa del mese
df_mensal = carregar_dados()

# 4. FESTIVI MOZAMBICO 2026
feriados_mz = {
    "01-01": "Ano Novo", "03-02": "Dia dos Heróis", "07-04": "Dia da Mulher",
    "01-05": "Dia do Trabalhador", "25-06": "Independência", "07-09": "Dia da Vitória",
    "25-09": "Dia das Forças Armadas", "04-10": "Dia da Paz", "25-12": "Natal"
}

# 5. AREA DI INSERIMENTO (Orario personalizzabile al minuto)
st.subheader("Registrar Dia de Trabalho")
col1, col2, col3 = st.columns(3)

with col1:
    data_sel = st.date_input("Data", datetime.now())

with col2:
    # Puoi scrivere l'ora esatta, es. 08:23
    h_in = st.time_input("Entrada", time(8, 0))

with col3:
    # Puoi scrivere l'ora esatta, es. 16:35
    h_out = st.time_input("Saída", time(16, 30))

# Controllo festivo automatico
dia_mes = data_sel.strftime("%d-%m")
nota_f = feriados_mz.get(dia_mes, "")
if nota_f:
    st.info(f"Feriado Moçambicano: {nota_f}")

# PULSANTE PER SALVARE NEL FOGLIO GOOGLE
if st.button("💾 Salvar Permanentemente"):
    t1 = datetime.combine(data_sel, h_in)
    t2 = datetime.combine(data_sel, h_out)
    total_h = (t2 - t1).total_seconds() / 3600

    if total_h <= 0:
        st.error("Erro: A hora de saída deve ser depois da entrada!")
    else:
        novo_registro = pd.DataFrame([{
            "Data": data_sel.strftime("%d/%m/%Y"),
            "Entrada": h_in.strftime("%H:%M"),
            "Saída": h_out.strftime("%H:%M"),
            "Horas": round(total_h, 2),
            "Obs": nota_f
        }])
        
        # Uniamo i vecchi dati con il nuovo e salviamo su Google
        df_aggiornato = pd.concat([df_mensal, novo_registro], ignore_index=True)
        conn.update(data=df_aggiornato)
        
        st.success(f"Dia {data_sel.strftime('%d/%m/%Y')} registrado com sucesso!")
        st.rerun()

# 6. LISTA COMPLETA DEL MESE (Visualizzazione storica)
st.divider()
st.subheader("📋 Histórico Mensal")

if not df_mensal.empty:
    # Mostra tutta la lista salvata finora
    st.dataframe(df_mensal, use_container_width=True, hide_index=True)
    
    # Calcolo totale ore accumulate
    st.info(f"**Total de horas no mês: {df_mensal['Horas'].sum():.2f} h**")

    # 7. FUNZIONE PER GENERARE L'EXCEL CON LE FIRME
    def criar_excel_assinatura(df):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Ponto', startrow=4)
            sheet = writer.sheets['Ponto']
            sheet['A1'] = "PARÓQUIA SS. TRINDADE - LIVRO DE PONTO"
            sheet['A2'] = f"Funcionária: Yolanda Facitela Clávio"
            sheet['A3'] = f"Pároco: Pe. Pasquale Peluso"
            
            # Linee per le firme in fondo all'Excel
            pos_firma = len(df) + 8
            sheet[f'A{pos_firma}'] = "__________________________"
            sheet[f'A{pos_firma+1}'] = "Assinatura Yolanda"
            sheet[f'D{pos_firma}'] = "__________________________"
            sheet[f'D{pos_firma+1}'] = "Assinatura Pe. Pasquale"
        return output.getvalue()

    # Bottone di download per Yolanda
    st.download_button(
        label="📥 Baixar Excel para Assinatura",
        data=criar_excel_assinatura(df_mensal),
        file_name=f"Livro_Ponto_{datetime.now().strftime('%m_%Y')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    
    # Opzione per resettare il database (nuovo mese)
    if st.checkbox("Mostrar opção para apagar histórico"):
        if st.button("⚠️ Apagar Tudo"):
            conn.update(data=pd.DataFrame(columns=["Data", "Entrada", "Saída", "Horas", "Obs"]))
            st.rerun()
else:
    st.write("Ainda não há dados salvos para este mês.")
