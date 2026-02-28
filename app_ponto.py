import streamlit as st
import pandas as pd
from datetime import datetime, time
import io

# 1. IMPOSTAZIONI PAGINA
st.set_page_config(page_title="Livro de Ponto", page_icon="📝")

# 2. INTESTAZIONE
st.title("📝 Livro de Ponto")
st.markdown("### **Paróquia SS. Trindade**")
st.write("**Pároco:** Pe. Pasquale Peluso")
st.write("**Funcionária:** Yolanda Facitela Clávio")
st.divider()

# 3. MEMORIA PER IL MESE (Salva i dati finché la pagina è aperta)
if 'registro_mensal' not in st.session_state:
    st.session_state['registro_mensal'] = pd.DataFrame(columns=["Data", "Entrada", "Saída", "Horas", "Obs"])

# 4. FESTIVI MOZAMBICO
feriados_mz = {
    "01-01": "Ano Novo", "03-02": "Dia dos Heróis", "07-04": "Dia da Mulher",
    "01-05": "Dia do Trabalhador", "25-06": "Independência", "07-09": "Dia da Vitória",
    "25-09": "Dia das Forças Armadas", "04-10": "Dia da Paz", "25-12": "Natal"
}

# 5. INSERIMENTO MANUALE
st.subheader("Registrar Dia de Trabalho")

# Layout a colonne per inserire i dati
col1, col2, col3 = st.columns(3)

with col1:
    # Qui Yolanda sceglie il giorno del mese
    data_sel = st.date_input("Escolha a Data", datetime.now())

with col2:
    # Qui scrive l'orario di entrata
    h_in = st.time_input("Hora de Entrada", time(8, 0))

with col3:
    # Qui scrive l'orario di uscita
    h_out = st.time_input("Hora de Saída", time(17, 0))

# Controllo festivi
dia_mes = data_sel.strftime("%d-%m")
obs_feriado = feriados_mz.get(dia_mes, "")
if obs_feriado:
    st.info(f"Nota: {obs_feriado} (Feriado)")

# Pulsante per aggiungere alla lista del mese
if st.button("Adicionar ao Relatório Mensal"):
    # Calcolo durata
    t1 = datetime.combine(data_sel, h_in)
    t2 = datetime.combine(data_sel, h_out)
    diff = t2 - t1
    total_h = diff.total_seconds() / 3600

    if total_h <= 0:
        st.error("A hora de saída deve ser maior que a entrada!")
    else:
        novo_registro = {
            "Data": data_sel.strftime("%d/%m/%Y"),
            "Entrada": h_in.strftime("%H:%M"),
            "Saída": h_out.strftime("%H:%M"),
            "Horas": round(total_h, 2),
            "Obs": obs_feriado
        }
        # Aggiunge alla tabella esistente
        st.session_state['registro_mensal'] = pd.concat([st.session_state['registro_mensal'], pd.DataFrame([novo_registro])], ignore_index=True)
        # Ordina per data automaticamente
        st.session_state['registro_mensal']['Data_dt'] = pd.to_datetime(st.session_state['registro_mensal']['Data'], format='%d/%m/%Y')
        st.session_state['registro_mensal'] = st.session_state['registro_mensal'].sort_values(by='Data_dt').drop(columns=['Data_dt'])
        st.success(f"Dia {data_sel.strftime('%d/%m')} adicionado!")

# 6. TABELLA RIASSUNTIVA DEL MESE
st.divider()
st.subheader("Resumo do Mês")
if not st.session_state['registro_mensal'].empty:
    st.dataframe(st.session_state['registro_mensal'], use_container_width=True)
    
    # Calcolo totale ore del mese
    total_mes = st.session_state['registro_mensal']['Horas'].sum()
    st.write(f"**Total acumulado no mês: {total_mes:.2f} horas**")

    # 7. GENERAZIONE EXCEL CON FIRME
    def prepare_excel(df):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Ponto', startrow=4)
            sheet = writer.sheets['Ponto']
            sheet['A1'] = "PARÓQUIA SS. TRINDADE - LIVRO DE PONTO"
            sheet['A2'] = f"Funcionária: Yolanda Facitela Clávio"
            sheet['A3'] = f"Pároco: Pe. Pasquale Peluso"
            
            # Linee per le firme in fondo
            last_row = len(df) + 7
            sheet[f'A{last_row}'] = "__________________________"
            sheet[f'A{last_row+1}'] = "Assinatura Yolanda"
            sheet[f'C{last_row}'] = "__________________________"
            sheet[f'C{last_row+1}'] = "Assinatura Pe. Pasquale"
        return output.getvalue()

    excel_data = prepare_excel(st.session_state['registro_mensal'])
    st.download_button(
        label="📥 Baixar Relatório Mensal (Excel)",
        data=excel_data,
        file_name=f"Livro_Ponto_Yolanda_{datetime.now().strftime('%m_%Y')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    
    # Bottone per pulire tutto e ricominciare il mese
    if st.button("Limpar tudo (Novo Mês)"):
        st.session_state['registro_mensal'] = pd.DataFrame(columns=["Data", "Entrada", "Saída", "Horas", "Obs"])
        st.rerun()
else:
    st.info("Ainda não há registros para este mês.")
