import streamlit as st
import pandas as pd
from datetime import datetime, time
import io

# 1. CONFIGURAZIONE DELLA PAGINA
st.set_page_config(page_title="Livro de Ponto", page_icon="📝", layout="centered")

# 2. INTESTAZIONE (PARROCCHIA E NOMI)
st.title("📝 Livro de Ponto")
st.markdown("### **Paróquia SS. Trindade**")
st.write("**Pároco:** Pe. Pasquale Peluso")
st.write("**Funcionária:** Yolanda Facitela Clávio")
st.divider()

# 3. MEMORIA TEMPORANEA (Per salvare i giorni inseriti durante la sessione)
if 'dados' not in st.session_state:
    st.session_state['dados'] = pd.DataFrame(columns=["Data", "Entrada", "Saída", "Horas", "Obs"])

# 4. FESTIVI MOZAMBICO
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
st.subheader("Registrar Horário de Hoje")

col1, col2 = st.columns(2)

with col1:
    data_hj = st.date_input("Data", datetime.now())
    h_entrada = st.time_input("Hora de Entrada", time(8, 0))

with col2:
    # Verifica se è festivo
    dia_mes = data_hj.strftime("%d-%m")
    nota_feriado = ""
    if dia_mes in feriados_mz:
        st.warning(f"Feriado: {feriados_mz[dia_mes]}")
        nota_feriado = feriados_mz[dia_mes]
    
    h_saida = st.time_input("Hora de Saída", time(17, 0))

# 6. CALCOLO DELLE ORE
dt_in = datetime.combine(data_hj, h_entrada)
dt_out = datetime.combine(data_hj, h_saida)
differenza = dt_out - dt_in
ore_totali = differenza.total_seconds() / 3600

if st.button("Adicionar ao Relatório"):
    if ore_totali <= 0:
        st.error("Erro: A hora de saída deve ser depois da entrada!")
    else:
        novo_dia = {
            "Data": data_hj.strftime("%d/%m/%Y"),
            "Entrada": h_entrada.strftime("%H:%M"),
            "Saída": h_saida.strftime("%H:%M"),
            "Horas": round(ore_totali, 2),
            "Obs": nota_feriado
        }
        # Aggiunge il giorno alla tabella
        st.session_state['dados'] = pd.concat([st.session_state['dados'], pd.DataFrame([novo_dia])], ignore_index=True)
        st.success("Dia adicionado com sucesso!")

# 7. VISUALIZZAZIONE TABELLA
st.divider()
st.subheader("Registros do Mês")
if not st.session_state['dados'].empty:
    st.table(st.session_state['dados'])
    
    # 8. FUNZIONE PER CREARE L'EXCEL CON LE FIRME
    def criar_excel(df):
        output = io.BytesIO()
        # Creiamo il file Excel
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Ponto_Mensal', startrow=2)
            
            # Accesso al foglio per aggiungere le intestazioni e le firme
            workbook = writer.book
            worksheet = writer.sheets['Ponto_Mensal']
            
            # Intestazione nel foglio Excel
            worksheet['A1'] = "LIVRO DE PONTO - PARÓQUIA SS. TRINDADE"
            worksheet['A2'] = f"Funcionária: Yolanda Facitela Clávio | Pároco: Pe. Pasquale Peluso"
            
            # Spazio per le firme in fondo
            ultima_linha = len(df) + 6
            worksheet[f'A{ultima_linha}'] = "__________________________"
            worksheet[f'A{ultima_linha+1}'] = "Assinatura da Funcionária"
            
            worksheet[f'D{ultima_linha}'] = "__________________________"
            worksheet[f'D{ultima_linha+1}'] = "Assinatura do Pároco"
            
        return output.getvalue()

    # Bottone di Download
    excel_file = criar_excel(st.session_state['dados'])
    st.download_button(
        label="📥 Baixar Excel para Assinatura",
        data=excel_file,
        file_name=f"Livro_de_Ponto_Yolanda_{datetime.now().strftime('%m_%Y')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("Nenhum dado inserido ainda.")