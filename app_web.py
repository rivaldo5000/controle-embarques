import streamlit as st
import pandas as pd
import sqlite3
from datetime import datetime
import io
import openpyxl
from openpyxl.styles import PatternFill

# Configuração da página Web
st.set_page_config(
    page_title="Controle de Embarques e Chegadas",
    page_icon="🚚",
    layout="wide"
)

DB_NAME = "controle_embarques.db"

# --- BANCO DE DADOS ---
def inicializar_banco():
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS embarques (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            cavalo TEXT,
            carreta TEXT,
            motorista TEXT,
            prev_chegada TEXT,
            loc_chegada TEXT,
            trava_chegada TEXT,
            chegada TEXT,
            status_prazo TEXT,
            prev_carreg TEXT,
            real_carreg TEXT,
            loc_carreg TEXT,
            trava_carreg TEXT,
            destino TEXT,
            frete TEXT,
            prazo TEXT
        )
    """)
    conn.commit()
    conn.close()

def converter_para_datetime(texto_data):
    if not texto_data or not str(texto_data).strip():
        return None
    
    texto_clean = str(texto_data).strip()
    formatos = [
        "%d/%m/%Y %H:%M", "%d/%m/%Y %H:%M:%S", "%d/%m/%Y",
        "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d"
    ]
    for fmt in formatos:
        try:
            return datetime.strptime(texto_clean, fmt)
        except ValueError:
            continue
    return "ERRO"

def calcular_status_prazo(prev_str, real_str):
    if not real_str or not str(real_str).strip():
        return "Pendente"
    if not prev_str or not str(prev_str).strip():
        return "Sem Previsão"

    dt_prev = converter_para_datetime(prev_str)
    dt_real = converter_para_datetime(real_str)

    if dt_prev == "ERRO" or dt_real == "ERRO":
        return "Erro Formato"

    return "Dentro do Prazo" if dt_real <= dt_prev else "Atraso"

inicializar_banco()

st.title("🚚 Controle de Embarques e Chegadas de Veículos")

# --- FORMULÁRIO DE CADASTRO / EDIÇÃO ---
with st.expander("📝 Formutário de Cadastro e Edição de Registros", expanded=True):
    # Opção para carregar registro existente para edição
    conn = sqlite3.connect(DB_NAME)
    df_lista = pd.read_sql_query("SELECT id, cavalo, carreta, motorista FROM embarques", conn)
    conn.close()

    opcoes_edicao = ["Novo Registro"] + [f"ID {row['id']} - Cavalo: {row['cavalo']} | Carreta: {row['carreta']}" for _, row in df_lista.iterrows()]
    selecao = st.selectbox("Selecione um registro para editar ou crie um novo:", opcoes_edicao)

    dados_atuais = {}
    id_selecionado = None

    if selecao != "Novo Registro":
        id_selecionado = int(selecao.split(" ")[1])
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM embarques WHERE id=?", (id_selecionado,))
        r = cursor.fetchone()
        conn.close()
        
        if r:
            dados_atuais = {
                "cavalo": r[1], "carreta": r[2], "motorista": r[3], "prev_chegada": r[4],
                "loc_chegada": r[5], "trava_chegada": r[6], "chegada": r[7], "prev_carreg": r[9],
                "real_carreg": r[10], "loc_carreg": r[11], "trava_carreg": r[12],
                "destino": r[13], "frete": r[14], "prazo": r[15]
            }

    with st.form("form_embarque", clear_on_submit=False):
        c1, c2, c3, c4 = st.columns(4)
        cavalo = c1.text_input("Placa Cavalo*", value=dados_atuais.get("cavalo", "")).upper()
        carreta = c2.text_input("Placa Carreta", value=dados_atuais.get("carreta", "")).upper()
        motorista = c3.text_input("Motorista", value=dados_atuais.get("motorista", ""))
        frete = c4.text_input("Valor Frete (R$)", value=dados_atuais.get("frete", ""))

        c5, c6, c7, c8 = st.columns(4)
        prev_chegada = c5.text_input("Prev. Chegada (Ex: 10/05/2026 14:00)", value=dados_atuais.get("prev_chegada", ""))
        loc_chegada = c6.text_input("Localizador (Chegada)", value=dados_atuais.get("loc_chegada", ""))
        trava_chegada = c7.text_input("Trava (Chegada)", value=dados_atuais.get("trava_chegada", ""))
        chegada_real = c8.text_input("Chegada Real (Ex: 10/05/2026 14:00)", value=dados_atuais.get("chegada", ""))

        c9, c10, c11, c12 = st.columns(4)
        prev_carreg = c9.text_input("Prev. Carregamento (Ex: 10/05/2026)", value=dados_atuais.get("prev_carreg", ""))
        real_carreg = c10.text_input("Data Real Carregamento (Ex: 10/05/2026)", value=dados_atuais.get("real_carreg", ""))
        loc_carreg = c11.text_input("Localizador (Carreg.)", value=dados_atuais.get("loc_carreg", ""))
        trava_carreg = c12.text_input("Trava (Carreg.)", value=dados_atuais.get("trava_carreg", ""))

        c13, c14 = st.columns(2)
        idx_destino = 0 if dados_atuais.get("destino", "") == "SEST/VDC" else 1
        destino = c13.selectbox("Destino da Carga", ["SEST/VDC", "SEST/VDC/SP"], index=idx_destino)
        prazo = c14.text_input("Prazo Pagamento", value=dados_atuais.get("prazo", ""))

        col_btn1, col_btn2 = st.columns([1, 5])
        salvar = col_btn1.form_submit_button("💾 Salvar Registro")

        if salvar:
            if not cavalo:
                st.error("A Placa do Cavalo é obrigatória!")
            else:
                status_prazo = calcular_status_prazo(prev_chegada, chegada_real)
                conn = sqlite3.connect(DB_NAME)
                cursor = conn.cursor()

                if id_selecionado is None:
                    cursor.execute("""
                        INSERT INTO embarques (
                            cavalo, carreta, motorista, prev_chegada, loc_chegada, trava_chegada,
                            chegada, status_prazo, prev_carreg, real_carreg, loc_carreg, trava_carreg,
                            destino, frete, prazo
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (
                        cavalo, carreta, motorista, prev_chegada, loc_chegada, trava_chegada,
                        chegada_real, status_prazo, prev_carreg, real_carreg, loc_carreg, trava_carreg,
                        destino, frete, prazo
                    ))
                    st.success("Registro adicionado com sucesso!")
                else:
                    cursor.execute("""
                        UPDATE embarques SET
                            cavalo=?, carreta=?, motorista=?, prev_chegada=?, loc_chegada=?, trava_chegada=?,
                            chegada=?, status_prazo=?, prev_carreg=?, real_carreg=?, loc_carreg=?, trava_carreg=?,
                            destino=?, frete=?, prazo=?
                        WHERE id=?
                    """, (
                        cavalo, carreta, motorista, prev_chegada, loc_chegada, trava_chegada,
                        chegada_real, status_prazo, prev_carreg, real_carreg, loc_carreg, trava_carreg,
                        destino, frete, prazo, id_selecionado
                    ))
                    st.success("Registro atualizado com sucesso!")

                conn.commit()
                conn.close()
                st.rerun()

# --- CONSULTA, FILTROS E TABELA ---
st.subheader("📊 Painel de Registros")

f1, f2, f3 = st.columns([2, 2, 2])
filtro_placa = f1.text_input("🔍 Filtrar por Placa (Cavalo ou Carreta)").strip().upper()
filtro_data = f2.text_input("📅 Filtrar por Data (Carregamento/Chegada)").strip()

conn = sqlite3.connect(DB_NAME)
df = pd.read_sql_query("SELECT * FROM embarques", conn)
conn.close()

if not df.empty:
    if filtro_placa:
        df = df[df["cavalo"].str.contains(filtro_placa, na=False) | df["carreta"].str.contains(filtro_placa, na=False)]
    if filtro_data:
        df = df[df["real_carreg"].str.contains(filtro_data, na=False) | df["prev_carreg"].str.contains(filtro_data, na=False)]

    # Aplicação de estilo de cores na tabela do site
    def aplicar_cores(row):
        chegada = row["chegada"]
        real_carreg = row["real_carreg"]
        
        cor = ""
        if not chegada:
            cor = "background-color: #ffcccc; color: black;"  # Vermelho
        elif chegada:
            cor = "background-color: #cce5ff; color: black;"  # Azul

        if not real_carreg and chegada:
            cor = "background-color: #fff3cd; color: black;"  # Amarelo
        elif real_carreg and chegada:
            cor = "background-color: #d4edda; color: black;"  # Verde

        return [cor] * len(row)

    # Exibição da tabela interativa
    st.dataframe(df.style.apply(aplicar_cores, axis=1), use_container_width=True, height=400)

    # --- GERADOR DE RELATÓRIO EXCEL ---
    def gerar_excel(data_frame):
        output = io.BytesIO()
        df_export = data_frame.copy()
        
        df_export.rename(columns={
            "id": "ID", "cavalo": "Placa Cavalo", "carreta": "Placa Carreta", 
            "motorista": "Motorista", "prev_chegada": "Prev. Chegada", 
            "loc_chegada": "Localizador (Chegada)", "trava_chegada": "Trava (Chegada)",
            "chegada": "Chegada Real", "status_prazo": "Status Prazo", 
            "prev_carreg": "Prev. Carregamento", "real_carreg": "Data Real Carregamento", 
            "loc_carreg": "Localizador (Carregamento)", "trava_carreg": "Trava (Carregamento)",
            "destino": "Destino", "frete": "Valor Frete", "prazo": "Prazo Pagamento"
        }, inplace=True)

        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_export.to_excel(writer, index=False, sheet_name='Embarques')
        
        output.seek(0)
        wb = openpyxl.load_workbook(output)
        ws = wb.active

        fill_vermelho = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
        fill_azul = PatternFill(start_color="CCE5FF", end_color="CCE5FF", fill_type="solid")
        fill_amarelo = PatternFill(start_color="FFF3CD", end_color="FFF3CD", fill_type="solid")
        fill_verde = PatternFill(start_color="D4EDDA", end_color="D4EDDA", fill_type="solid")

        for row in range(2, ws.max_row + 1):
            chegada_val = ws.cell(row=row, column=8).value
            real_carreg_val = ws.cell(row=row, column=11).value

            fill_aplicar = None
            if not chegada_val:
                fill_aplicar = fill_vermelho
            elif chegada_val:
                fill_aplicar = fill_azul

            if not real_carreg_val and chegada_val:
                fill_aplicar = fill_amarelo
            elif real_carreg_val and chegada_val:
                fill_aplicar = fill_verde

            if fill_aplicar:
                for col in range(1, ws.max_column + 1):
                    ws.cell(row=row, column=col).fill = fill_aplicar

        out_final = io.BytesIO()
        wb.save(out_final)
        out_final.seek(0)
        return out_final

    excel_data = gerar_excel(df)
    st.download_button(
        label="📥 Baixar Relatório Formatado em Excel (.xlsx)",
        data=excel_data,
        file_name=f"relatorio_embarques_{datetime.now().strftime('%d_%m_%Y')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("Nenhum registro encontrado no banco de dados.")