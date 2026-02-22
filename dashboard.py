import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import calendar
import os
import sys
import urllib.parse
import openpyxl 
import warnings
import time
from filelock import FileLock, Timeout
import win32com.client
import pythoncom

warnings.simplefilter(action='ignore', category=UserWarning)

st.set_page_config(page_title="Gestão de Projetos", layout="wide")

# =====================================================================
# --- CSS CUSTOMIZADO E ESTILIZAÇÃO (com borda inferior nas abas) ---
# =====================================================================
st.markdown("""
<style>
    /* Sidebar */
    section[data-testid="stSidebar"] { background-color: #FF4F00; }
    section[data-testid="stSidebar"] h1, section[data-testid="stSidebar"] h2, 
    section[data-testid="stSidebar"] h3, section[data-testid="stSidebar"] label, 
    section[data-testid="stSidebar"] div[data-testid="stMarkdownContainer"] p,
    section[data-testid="stSidebar"] .stMultiSelect div[data-baseweb="select"] > div { color: white !important; }
    section[data-testid="stSidebar"] div[data-baseweb="popover"] div,
    section[data-testid="stSidebar"] div[data-baseweb="select"] span { color: black !important; }
    
    /* Botões da Sidebar */
    section[data-testid="stSidebar"] .stButton button {
        background-color: transparent !important;
        color: white !important;
        border: 1px solid white !important;
        transition: 0.3s;
    }
    section[data-testid="stSidebar"] .stButton button:hover {
        background-color: white !important;
        color: #FF4F00 !important;
        border-color: white !important;
    }
    section[data-testid="stSidebar"] .stButton button p {
        color: inherit !important; 
    }

    /* --- ESTILO PARA ABAS (st.radio ou st.segmented_control) COM BORDA INFERIOR --- */
    /* Para st.radio */
    div.row-widget.stRadio > div[role="radiogroup"] {
        display: flex !important;
        flex-direction: row !important;
        gap: 0.5rem !important;
        background-color: #f0f2f6 !important;
        padding: 0.5rem 0.5rem 0 0.5rem !important;
        border-radius: 2rem 2rem 0 0 !important;
        justify-content: center !important;
        margin-bottom: 1.5rem !important;
    }
    div.row-widget.stRadio > div[role="radiogroup"] label {
        background-color: transparent !important;
        padding: 0.5rem 1.5rem !important;
        border-radius: 2rem 2rem 0 0 !important;
        font-weight: 600 !important;
        color: #333 !important;
        transition: all 0.2s !important;
        margin: 0 !important;
        display: inline-flex !important;
        align-items: center !important;
    }
    div.row-widget.stRadio > div[role="radiogroup"] label > div:first-child {
        display: none !important;
    }
    div.row-widget.stRadio > div[role="radiogroup"] label:hover {
        background-color: rgba(255, 79, 0, 0.1) !important;
    }
    div.row-widget.stRadio > div[role="radiogroup"] label[data-baseweb="radio"] input:checked + * {
        background-color: #FF4F00 !important;
        color: white !important;
        border-radius: 2rem 2rem 0 0 !important;
    }

    /* Para st.segmented_control (quando disponível) */
    div[data-testid="stSegmentedControl"] {
        margin-bottom: 1.5rem !important;
        padding-bottom: 0 !important;
    }
    div[data-testid="stSegmentedControl"] div[data-baseweb="segmented-control"] {
        background-color: #f0f2f6 !important;
        border-radius: 2rem 2rem 0 0 !important;
        padding: 0.25rem 0.25rem 0 0.25rem !important;
    }
    div[data-testid="stSegmentedControl"] button {
        border-radius: 2rem 2rem 0 0 !important;
        border: none !important;
        font-weight: 600 !important;
        padding: 0.5rem 1.5rem !important;
    }
    div[data-testid="stSegmentedControl"] button[aria-selected="true"] {
        background-color: #FF4F00 !important;
        color: white !important;
    }

    /* Estilo para os botões de navegação da tela de projeto */
    div[data-testid="column"] .stButton button {
        width: 100%;
        font-size: 1rem;
        padding: 0.5rem 1rem;
        border-radius: 2rem;
        border: 1px solid #FF4F00;
        background-color: white;
        color: #FF4F00;
        transition: all 0.2s;
    }
    div[data-testid="column"] .stButton button:hover {
        background-color: #FF4F00;
        color: white;
        border-color: #FF4F00;
    }
    /* Alinhamento vertical dos botões */
    div[data-testid="column"] {
        display: flex;
        align-items: center;
        justify-content: center;
    }
    /* Pequeno espaçamento superior */
    div.sticky-nav {
        margin-top: 1rem;
        margin-bottom: 1rem;
    }
</style>
""", unsafe_allow_html=True)

if 'modo_edicao' not in st.session_state:
    st.session_state.modo_edicao = False
if 'projeto_ativo_state' not in st.session_state:
    st.session_state.projeto_ativo_state = "-- Visão Geral (Dashboard) --"
if 'aba_ativa' not in st.session_state:
    st.session_state.aba_ativa = "Visão Geral"

params = st.query_params
if "projeto" in params:
    st.session_state.projeto_ativo_state = params["projeto"]
    st.session_state.modo_edicao = False 
    st.query_params.clear()

# =====================================================================
# --- FUNÇÕES DE BACKEND (inalteradas) ---
# =====================================================================
ARQUIVO_ENTRADA = 'CEPA Ações.xlsm'
ARQUIVO_SAIDA = 'Base_Dados_Consolidada_CEPA.xlsx'

def obter_caminho(nome_arquivo):
    if getattr(sys, 'frozen', False):
        basedir = os.path.dirname(sys.executable)
    else:
        basedir = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(basedir, nome_arquivo)

def formatar_data(d):
    if pd.notnull(d) and hasattr(d, 'year'):
        return datetime(d.year, d.month, d.day)
    return None

@st.cache_data(ttl=2) 
def carregar_dados():
    arquivo = obter_caminho(ARQUIVO_SAIDA)
    try:
        df = pd.read_excel(arquivo)
        cols_data = ['Prazo_Projeto', 'Inicio_Projeto', 'Data_Conclusao_Projeto', 'Prazo_Tarefa', 'Data_Conclusao_Tarefa']
        for col in cols_data:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors='coerce')
        
        df['Percentual_Conclusao'] = pd.to_numeric(df.get('Percentual_Conclusao', 0.0), errors='coerce').fillna(0)
        
        if 'Inicio_Projeto' in df.columns:
            df['Ano_Inicio'] = df['Inicio_Projeto'].dt.year.fillna(0).astype(int)
            df['Mes_Inicio'] = df['Inicio_Projeto'].dt.month_name()
            
        cols_str = ['Projeto', 'Área', 'Status_Projeto', 'Concluido', 'Status_Prazo_Tarefa', 'Classe', 'Objetivo', 'Tipo_Obs', 'Texto_Obs', 'Resumo', 'Passos_Criticos', 'Riscos']
        for col in cols_str:
             if col in df.columns:
                df[col] = df[col].astype(str).str.strip().replace('nan', '')
                
        df['Classe'] = df.get('Classe', 'Média').replace('', 'Média')
        
        def calcular_dias(row):
            inicio = row['Inicio_Projeto']
            prazo = row['Prazo_Projeto']
            conclusao = row['Data_Conclusao_Projeto']
            status = str(row['Status_Projeto']).lower()
            if pd.isnull(inicio) or pd.isnull(prazo): return 0, 0
            dias_planejados = (prazo - inicio).days
            dias_atraso = 0
            hoje = datetime.today()
            if pd.notnull(conclusao):
                if conclusao > prazo: dias_atraso = (conclusao - prazo).days
            else:
                if 'hold' not in status and 'cancelado' not in status and 'parado' not in status:
                    if hoje > prazo: dias_atraso = (hoje - prazo).days
            return dias_planejados, dias_atraso

        if 'Inicio_Projeto' in df.columns and 'Prazo_Projeto' in df.columns:
             df[['Dias_Planejados', 'Dias_Atraso']] = df.apply(lambda row: pd.Series(calcular_dias(row)), axis=1)

        return df
    except Exception as e:
        return pd.DataFrame()

def reconstruir_banco_de_dados():
    arquivo_in = obter_caminho(ARQUIVO_ENTRADA)
    arquivo_out = obter_caminho(ARQUIVO_SAIDA)

    def encontrar_aba_correspondente(nome_projeto_lista, todas_abas_excel):
        if nome_projeto_lista in todas_abas_excel: return nome_projeto_lista
        parte_nome = str(nome_projeto_lista)[:30].strip()
        for aba in todas_abas_excel:
            if aba.startswith(parte_nome): return aba
        nome_limpo = str(nome_projeto_lista)
        for char in ['/', '\\', '?', '*', '[', ']', ':']:
            nome_limpo = nome_limpo.replace(char, '-')
        parte_limpo = nome_limpo[:30].strip()
        for aba in todas_abas_excel:
            if aba.startswith(parte_limpo): return aba
        return None

    def safe_iloc(df, row, col, default=""):
        try:
            val = df.iloc[row, col]
            return str(val).strip() if pd.notnull(val) else default
        except: return default

    try:
        xls = pd.ExcelFile(arquivo_in, engine='openpyxl')
        df_projetos = pd.read_excel(xls, sheet_name='Projetos')
        df_projetos.columns = [str(c).strip() for c in df_projetos.columns]
        
        lista_projetos = []
        if 'Título' in df_projetos.columns:
            lista_projetos = df_projetos['Título'].dropna().unique().tolist()
            
        if not lista_projetos:
            for aba in xls.sheet_names:
                if aba not in ['Projetos', 'Resumo', 'Dashboard', 'Base', 'Config']:
                    lista_projetos.append(aba)

        if not lista_projetos:
            return False, "Nenhum projeto encontrado."

        dados_consolidados = []
        
        for projeto_nome in lista_projetos:
            aba_real = encontrar_aba_correspondente(projeto_nome, xls.sheet_names)
            if aba_real:
                try:
                    df_aba = pd.read_excel(xls, sheet_name=aba_real, header=None)
                    
                    status = safe_iloc(df_aba, 0, 8)  
                    area = safe_iloc(df_aba, 1, 8) 
                    prazo_proj = df_aba.iloc[2, 1]    
                    inicio_proj = df_aba.iloc[2, 5]   
                    classe_proj = safe_iloc(df_aba, 3, 1) or "Média" 
                    
                    objetivo = safe_iloc(df_aba, 1, 1)
                    tipo_obs = safe_iloc(df_aba, 2, 6)
                    texto_obs = safe_iloc(df_aba, 2, 7)
                    resumo = safe_iloc(df_aba, 5, 0)
                    passos_criticos = safe_iloc(df_aba, 11, 0)
                    riscos = safe_iloc(df_aba, 14, 0)

                    val_d3 = df_aba.iloc[2, 3] 
                    val_i9 = df_aba.iloc[8, 8] 
                    
                    data_conclusao_proj = None 
                    percentual = 0.0
                    if isinstance(val_i9, (int, float)): percentual = float(val_i9)
                    
                    is_concluido = False
                    if str(status).strip().lower() == 'concluído':
                        is_concluido = True
                        percentual = 1.0
                    elif percentual >= 0.99:
                        is_concluido = True
                    
                    if is_concluido: data_conclusao_proj = val_d3

                    linha_inicio = 4
                    for i, valor in enumerate(df_aba.iloc[:, 10]):
                        if isinstance(valor, str) and "Planejamento de Atividades" in valor:
                            linha_inicio = i + 1
                            break

                    tarefas = df_aba.iloc[linha_inicio:, [10, 11, 12, 13, 14]].copy()
                    tarefas.columns = ['Descricao', 'Concluido', 'Prazo_Tarefa', 'Data_Conclusao_Tarefa', 'Status_Prazo_Tarefa']
                    tarefas = tarefas[tarefas['Descricao'].notna()].copy()

                    if not tarefas.empty:
                        tarefas['Projeto'] = str(projeto_nome)
                        tarefas['Status_Projeto'] = status
                        tarefas['Área'] = area
                        tarefas['Prazo_Projeto'] = prazo_proj
                        tarefas['Inicio_Projeto'] = inicio_proj
                        tarefas['Data_Conclusao_Projeto'] = data_conclusao_proj
                        tarefas['Classe'] = classe_proj
                        tarefas['Percentual_Conclusao'] = percentual
                        tarefas['Objetivo'] = objetivo
                        tarefas['Tipo_Obs'] = tipo_obs
                        tarefas['Texto_Obs'] = texto_obs
                        tarefas['Resumo'] = resumo
                        tarefas['Passos_Criticos'] = passos_criticos
                        tarefas['Riscos'] = riscos
                        
                        tarefas['Concluido'] = tarefas['Concluido'].apply(lambda x: 'Sim' if x in [1, True, 'True', 'Sim'] else 'Não')
                        dados_consolidados.append(tarefas)
                except Exception as e:
                    pass 
        
        if dados_consolidados:
            df_final = pd.concat(dados_consolidados, ignore_index=True)
            df_final.to_excel(arquivo_out, index=False)
            return True, "Banco atualizado com sucesso."
        else:
            return False, "Nenhum dado válido de tarefa foi encontrado."

    except Exception as e:
        return False, f"Erro ao reconstruir base: {e}"

def salvar_alteracoes_no_excel(projeto_nome, dados_editados, status_tarefas, concluir_projeto=False, nova_tarefa_desc="", nova_tarefa_prazo=None):
    arquivo_abs = os.path.abspath(obter_caminho(ARQUIVO_ENTRADA))
    caminho_lock = obter_caminho(ARQUIVO_ENTRADA + ".lock")
    lock = FileLock(caminho_lock, timeout=15)

    try:
        with open(arquivo_abs, 'a'): pass
    except IOError:
        return False, "O arquivo Excel está ABERTO em outro local. Feche-o para salvar."

    def limpar_caracteres_aba(nome):
        n = str(nome)
        for char in ['/', '\\', '?', '*', '[', ']', ':']:
            n = n.replace(char, '-')
        return n[:30].strip()

    try:
        with lock:
            pythoncom.CoInitialize() 
            try:
                excel = win32com.client.DispatchEx("Excel.Application")
                excel.Visible = False
                excel.DisplayAlerts = False
            except Exception as e:
                return False, f"Erro ao acionar MS Excel: {e}"

            try:
                wb = excel.Workbooks.Open(arquivo_abs)
                if wb.ReadOnly:
                    wb.Close(SaveChanges=False)
                    excel.Quit()
                    return False, "O arquivo Excel está bloqueado."

                nome_aba = str(projeto_nome)[:30].strip()
                ws = None
                for sheet in wb.Sheets:
                    if sheet.Name.startswith(nome_aba) or sheet.Name.startswith(limpar_caracteres_aba(projeto_nome)):
                        ws = sheet; break
                
                if not ws:
                    wb.Close(SaveChanges=False)
                    return False, "Aba não encontrada para edição."

                ws.Range("B2").Value = dados_editados['objetivo']
                ws.Range("A6").Value = dados_editados['resumo']
                ws.Range("A12").Value = dados_editados['passos_criticos']
                ws.Range("A15").Value = dados_editados['riscos']
                ws.Range("H3").Value = dados_editados['texto_obs'] 
                
                if dados_editados['inicio_proj']: ws.Range("F3").Value = formatar_data(dados_editados['inicio_proj'])
                if dados_editados['prazo_proj']: ws.Range("B3").Value = formatar_data(dados_editados['prazo_proj'])

                if concluir_projeto:
                    ws.Range("I1").Value = "Concluído"
                    ws.Range("D3").Value = formatar_data(datetime.today())
                    for k in status_tarefas.keys(): 
                        status_tarefas[k]['concluido'] = True
                        status_tarefas[k]['data_conclusao'] = datetime.today()
                    
                    try:
                        ws_proj = wb.Sheets("Projetos")
                        tabela = None
                        for tbl in ws_proj.ListObjects:
                            if tbl.Name == "Atividades": tabela = tbl; break
                        
                        if tabela:
                            col_titulo = tabela.ListColumns("Título").Index
                            col_data = None
                            try: col_data = tabela.ListColumns("Data de Conclusão").Index
                            except: pass
                            
                            if col_data:
                                for list_row in tabela.ListRows:
                                    if str(list_row.Range(1, col_titulo).Value).strip() == str(projeto_nome).strip():
                                        list_row.Range(1, col_data).Value = formatar_data(datetime.today())
                                        break
                    except Exception as err:
                        pass

                linha_inicio = 4
                for row in range(1, 30):
                    val = ws.Cells(row, 11).Value
                    if val and "Planejamento de Atividades" in str(val):
                        linha_inicio = row + 1; break

                max_row = ws.UsedRange.Rows.Count
                linha_vazia_tarefa = linha_inicio

                for row in range(linha_inicio, max_row + 10):
                    desc = ws.Cells(row, 11).Value
                    if not desc: 
                        if linha_vazia_tarefa <= row: linha_vazia_tarefa = row
                        continue
                        
                    linha_vazia_tarefa = row + 1 
                    desc_str = str(desc).strip()
                    
                    if desc_str in status_tarefas:
                        dados_t = status_tarefas[desc_str]
                        ws.Cells(row, 12).Value = bool(dados_t['concluido'])
                        if dados_t['prazo']: ws.Cells(row, 13).Value = formatar_data(dados_t['prazo'])
                        
                        data_conc = dados_t['data_conclusao']
                        if dados_t['concluido']:
                            if pd.notnull(data_conc): ws.Cells(row, 14).Value = formatar_data(data_conc)
                            else: ws.Cells(row, 14).Value = formatar_data(datetime.today())
                        else:
                            ws.Cells(row, 14).Value = None 

                if nova_tarefa_desc:
                    ws.Cells(linha_vazia_tarefa, 11).Value = nova_tarefa_desc
                    if nova_tarefa_prazo: 
                        ws.Cells(linha_vazia_tarefa, 13).Value = formatar_data(nova_tarefa_prazo)

                wb.Save()
                wb.Close(SaveChanges=True)

            finally:
                excel.Quit()
                pythoncom.CoUninitialize() 
                
        sucesso_bd, msg = reconstruir_banco_de_dados()
        carregar_dados.clear()
        if sucesso_bd:
            return True, "Alterações salvas e painel sincronizado!"
        else:
            return True, f"Salvo com sucesso, mas ocorreu aviso ao recarregar tela: {msg}"
            
    except Timeout:
        return False, "O sistema está processando outra edição no momento. Aguarde alguns segundos."
    except Exception as e:
        return False, f"Erro interno ao salvar: {e}"

def executar_criacao_projeto(titulo, area, prazo_proj, objetivo, classe_proj, obs, resumo, passos, riscos, df_tarefas):
    arquivo_abs = os.path.abspath(obter_caminho(ARQUIVO_ENTRADA))
    caminho_lock = obter_caminho(ARQUIVO_ENTRADA + ".lock")
    lock = FileLock(caminho_lock, timeout=20)

    try:
        with open(arquivo_abs, 'a'): pass
    except IOError:
        return False, "O arquivo Excel está ABERTO. Feche-o para criar o projeto."

    try:
        with lock:
            pythoncom.CoInitialize()
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            try:
                wb = excel.Workbooks.Open(arquivo_abs)
                ws_proj = wb.Sheets("Projetos")

                tabela = None
                for tbl in ws_proj.ListObjects:
                    if tbl.Name == "Atividades": tabela = tbl; break
                
                if tabela:
                    nova_linha = tabela.ListRows.Add()
                    col_titulo = tabela.ListColumns("Título").Index
                    nova_linha.Range(1, col_titulo).Value = titulo
                else:
                    return False, "Tabela 'Atividades' não foi localizada na aba 'Projetos'."
                
                nome_macro = f"'{wb.Name}'!AtualizarECriarProjetos"
                try:
                    excel.Application.Run(nome_macro)
                except Exception as err_macro:
                    pass 
                
                nome_aba = str(titulo)[:30].strip()
                ws_nova = None
                for sheet in wb.Sheets:
                    if sheet.Name.startswith(nome_aba):
                        ws_nova = sheet; break
                        
                if ws_nova:
                    ws_nova.Range("I1").Value = "Em Andamento"
                    ws_nova.Range("I2").Value = area
                    ws_nova.Range("F3").Value = formatar_data(datetime.today())
                    if prazo_proj: ws_nova.Range("B3").Value = formatar_data(prazo_proj)
                    
                    ws_nova.Range("B2").Value = objetivo
                    ws_nova.Range("B4").Value = classe_proj
                    ws_nova.Range("G3").Value = "Observações" 
                    ws_nova.Range("H3").Value = obs
                    
                    ws_nova.Range("A6").Value = resumo
                    ws_nova.Range("A12").Value = passos
                    ws_nova.Range("A15").Value = riscos
                    
                    linha_inicio = 4
                    for row in range(1, 30):
                        val = ws_nova.Cells(row, 11).Value
                        if val and "Planejamento de Atividades" in str(val):
                            linha_inicio = row + 1; break
                            
                    linha_atual = linha_inicio
                    for _, row_data in df_tarefas.iterrows():
                        tarefa_desc = str(row_data["Tarefa"]).strip()
                        if tarefa_desc:
                            ws_nova.Cells(linha_atual, 11).Value = tarefa_desc
                            ws_nova.Cells(linha_atual, 12).Value = False 
                            if pd.notnull(row_data["Prazo"]):
                                ws_nova.Cells(linha_atual, 13).Value = formatar_data(row_data["Prazo"])
                            linha_atual += 1

                wb.Save()
                wb.Close(SaveChanges=True)
            finally:
                excel.Quit()
                pythoncom.CoUninitialize()
        
        sucesso_bd, msg = reconstruir_banco_de_dados()
        carregar_dados.clear()
        return True, "Projeto criado, populado e banco atualizado!"
    except Exception as e:
        return False, f"Erro na criação: {e}"

# =====================================================================
# --- POP-UP (MODAL) DE CRIAÇÃO ---
# =====================================================================
@st.dialog("Formulário de Criação de Projeto", width="large")
def modal_novo_projeto():
    st.write("Preencha os dados abaixo para inicializar a aba perfeitamente.")
    
    col1, col2 = st.columns(2)
    titulo_novo = col1.text_input("Título do Projeto *", max_chars=30)
    area_novo = col2.text_input("Área(s) *", help="Separe por ponto-e-vírgula se for mais de um.")
    
    col3, col4 = st.columns(2)
    classe_novo = col3.selectbox("Prioridade", ["Alta", "Média", "Baixa"], index=1)
    prazo_novo = col4.date_input("Prazo Final *")
    
    st.write("---")
    st.markdown("**Detalhamento Principal**")
    objetivo_novo = st.text_area("Objetivo do Projeto", height=68)
    
    c_text1, c_text2 = st.columns(2)
    resumo_novo = c_text1.text_area("Resumo do Projeto", height=100)
    obs_novo = c_text2.text_area("Observações Importantes", height=100)
    
    c_passos, c_riscos = st.columns(2)
    passos_novo = c_passos.text_area("Passos Críticos", height=100)
    riscos_novo = c_riscos.text_area("Riscos Mapeados", height=100)
    
    st.write("---")
    st.markdown("**Inserção Rápida de Tarefas**")
    st.caption("Aperte o botão + abaixo para adicionar quantas tarefas quiser ao novo projeto.")
    
    df_tarefas_vazio = pd.DataFrame([{"Tarefa": "", "Prazo": datetime.today().date()}])
    df_tarefas_editado = st.data_editor(
        df_tarefas_vazio, 
        num_rows="dynamic", 
        use_container_width=True,
        column_config={
            "Tarefa": st.column_config.TextColumn("Descrição da Tarefa", required=True),
            "Prazo": st.column_config.DateColumn("Prazo Limite", format="DD/MM/YYYY")
        }
    )
    
    if st.button("Gravar no Excel e Gerar Aba", type="primary", use_container_width=True):
        if not titulo_novo:
            st.error("O título é obrigatório!")
            return
            
        tarefas_validas = df_tarefas_editado[df_tarefas_editado["Tarefa"].str.strip() != ""]
        if tarefas_validas.empty:
            st.error("Insira pelo menos 1 tarefa válida com descrição!")
            return
            
        with st.spinner("Acionando motor do Excel em segundo plano... Isso pode levar alguns segundos."):
            sucesso, msg = executar_criacao_projeto(
                titulo_novo, area_novo, prazo_novo, objetivo_novo, classe_novo, obs_novo, resumo_novo, passos_novo, riscos_novo, tarefas_validas
            )
            if sucesso:
                st.success(msg)
                time.sleep(1)
                st.rerun()
            else:
                st.error(msg)

# =====================================================================
# --- INTERFACE PRINCIPAL ---
# =====================================================================

# --- NOVO: AUTO-INICIALIZAÇÃO DO BANCO DE DADOS ---
if 'banco_inicializado' not in st.session_state:
    caminho_entrada = obter_caminho(ARQUIVO_ENTRADA)
    
    # Só tenta reconstruir se a planilha macro (CEPA Ações.xlsm) existir na pasta
    if os.path.exists(caminho_entrada):
        with st.spinner("Inicializando e Sincronizando banco de dados..."):
            reconstruir_banco_de_dados()
            carregar_dados.clear() # Limpa a memória para garantir a leitura fresca
            
    # Marca que já foi inicializado para não travar o app nos próximos cliques
    st.session_state.banco_inicializado = True
# --------------------------------------------------

df = carregar_dados()

if not df.empty:
    
    url_logo = "https://raw.githubusercontent.com/Agrivalle-BioDigital/Logo_AGVL/main/AF_logo_agrivalle_novo_Branco._rgbai-removebg.png"
    st.sidebar.markdown(f"""<div style="display: flex; justify-content: center; width: 100%; margin-bottom: 20px;"><img src="{url_logo}" width="180"></div>""", unsafe_allow_html=True)
    
    todos_projetos_lista = ["-- Visão Geral (Dashboard) --"] + sorted(df['Projeto'].unique().tolist())
    
    def reset_edicao():
        st.session_state.modo_edicao = False
        
    # Garante que o valor no session_state seja uma opção válida da lista atual
    if st.session_state.projeto_ativo_state not in todos_projetos_lista:
        st.session_state.projeto_ativo_state = "-- Visão Geral (Dashboard) --"

    # Usando o 'key', o Streamlit lê e escreve o valor diretamente no st.session_state
    st.sidebar.selectbox(
        "Detalhar Projeto:", 
        todos_projetos_lista, 
        key="projeto_ativo_state", 
        on_change=reset_edicao
    )

    st.sidebar.markdown("<br>", unsafe_allow_html=True)
    if st.sidebar.button("Criar Novo Projeto", use_container_width=True, type="primary"):
        modal_novo_projeto()

    st.sidebar.markdown("<br><hr style='border:1px solid white; opacity:0.3;'><br>", unsafe_allow_html=True)
    st.sidebar.header("Filtros Globais")
    
    todos_areas = df['Área'].dropna().unique().tolist()
    set_areas = set()
    for item in todos_areas:
        for p in item.split(';'): set_areas.add(p.strip())
    
    cli_sel = st.sidebar.selectbox("Área", ['Todos'] + sorted(list(set_areas)))
    sts_sel = st.sidebar.selectbox("Status Projeto", ['Todos'] + sorted(df['Status_Projeto'].unique().tolist()))
    
    if 'Ano_Inicio' in df.columns:
        anos = sorted([x for x in df['Ano_Inicio'].unique() if x > 0])
        ano_sel = st.sidebar.multiselect("Ano Início", anos, default=anos)
    else: ano_sel = []

    st.sidebar.markdown("<br><hr style='border:1px solid white; opacity:0.3;'>", unsafe_allow_html=True)
    if st.sidebar.button("Sincronizar com Excel Original", use_container_width=True, type="secondary"):
        with st.spinner("Lendo Excel e atualizando dashboard..."):
            sucesso, msg = reconstruir_banco_de_dados()
            if sucesso:
                carregar_dados.clear()
                st.sidebar.success("Base atualizada com sucesso!")
                time.sleep(1)
                st.rerun()
            else:
                st.sidebar.error(f"Erro: {msg}")

    df_f = df.copy()
    if cli_sel != 'Todos': df_f = df_f[df_f['Área'].str.contains(cli_sel, regex=False)]
    if sts_sel != 'Todos': df_f = df_f[df_f['Status_Projeto'] == sts_sel]
    if ano_sel: df_f = df_f[df_f['Ano_Inicio'].isin(ano_sel)]

    # =====================================================================
    # --- TELA 2: VISÃO DETALHADA DO PROJETO (com barra de navegação simplificada) ---
    # =====================================================================
    if st.session_state.projeto_ativo_state != "-- Visão Geral (Dashboard) --":
        df_proj = df[df['Projeto'] == st.session_state.projeto_ativo_state].copy()
        
        if not df_proj.empty:
            info = df_proj.iloc[0] 
            
            # --- BARRA DE NAVEGAÇÃO SIMPLIFICADA (apenas dois botões) ---
            st.markdown('<div class="sticky-nav">', unsafe_allow_html=True)
            col_esq, col_dir = st.columns(2)
            with col_esq:
                # 1. Criamos a função de callback
                def ir_para_inicio():
                    st.session_state.projeto_ativo_state = "-- Visão Geral (Dashboard) --"
                    st.session_state.modo_edicao = False
                
                # 2. Atrelamos a função ao botão usando on_click
                # O Streamlit fará o rerun automaticamente, não precisamos mais do st.rerun()
                st.button("Início", use_container_width=True, on_click=ir_para_inicio)
            with col_dir:
                icone = "Cancelar Edição" if st.session_state.modo_edicao else "Editar"
                if st.button(icone, type="secondary" if st.session_state.modo_edicao else "primary", use_container_width=True):
                    st.session_state.modo_edicao = not st.session_state.modo_edicao
                    st.rerun()
            st.markdown('</div>', unsafe_allow_html=True)

            st.markdown(f"<h1 style='color: #FF4F00; margin-bottom: 0rem;'>{info['Projeto']}</h1>", unsafe_allow_html=True)
            area_display = info.get('Área', 'Não informado')
            st.markdown(f"<h4 style='color: #666; margin-top: 0rem; font-weight: 500;'>Área(s): {area_display}</h4>", unsafe_allow_html=True)
            st.write("")
            
            hoje = datetime.today()
            dias_restantes = 0
            duracao_total = 0
            if pd.notnull(info.get('Prazo_Projeto')) and pd.notnull(info.get('Inicio_Projeto')):
                duracao_total = (info['Prazo_Projeto'] - info['Inicio_Projeto']).days
                dias_restantes = (info['Prazo_Projeto'] - hoje).days

            colA, colB, colC, colD = st.columns(4)
            colA.metric("Status Atual", info.get('Status_Projeto', 'N/A'))
            colB.metric("Prioridade", info.get('Classe', 'Média'))
            colC.metric("Duração Planejada", f"{duracao_total} dias")
            if dias_restantes < 0:
                colD.metric("Dias Restantes", f"Atrasado {-dias_restantes} dias", delta_color="inverse")
            else:
                colD.metric("Dias Restantes", f"{dias_restantes} dias")

            pct = info.get('Percentual_Conclusao', 0)
            st.markdown(f"**Progresso do Projeto: {pct*100:.0f}%**")
            st.progress(pct)
            st.write("---")

            if not st.session_state.modo_edicao:
                colE, colF = st.columns(2)
                with colE:
                    st.subheader("Objetivo")
                    st.info(info.get('Objetivo', 'Nenhum objetivo informado.') or 'Nenhum objetivo informado.')
                    st.subheader("Resumo do Projeto")
                    st.write(info.get('Resumo', 'Nenhum resumo informado.') or 'Nenhum resumo informado.')
                with colF:
                    titulo_obs = info.get('Tipo_Obs', 'Observações')
                    st.subheader(f"{titulo_obs}")
                    texto_obs = info.get('Texto_Obs', 'Sem observações.')
                    if "Justificativa" in titulo_obs and texto_obs: st.error(texto_obs)
                    else: st.warning(texto_obs or 'Sem observações.')
                    st.subheader("Riscos")
                    st.write(info.get('Riscos', 'Nenhum risco mapeado.') or 'Nenhum risco mapeado.')

                st.subheader("Passos Críticos")
                st.write(info.get('Passos_Criticos', 'Nenhum passo crítico informado.') or 'Nenhum passo crítico informado.')
                st.write("---")
                
                st.subheader("Tarefas do Projeto")
                cols_tarefa = ['Descricao', 'Concluido', 'Prazo_Tarefa', 'Data_Conclusao_Tarefa', 'Status_Prazo_Tarefa']
                df_tarefas_exibir = df_proj[cols_tarefa].copy()
                df_tarefas_exibir = df_tarefas_exibir[df_tarefas_exibir['Descricao'] != '']
                
                st.dataframe(
                    df_tarefas_exibir,
                    use_container_width=True,
                    hide_index=True,
                    column_config={
                        "Descricao": st.column_config.TextColumn("Tarefa"),
                        "Concluido": st.column_config.TextColumn("Feito?"),
                        "Prazo_Tarefa": st.column_config.DateColumn("Prazo", format="DD/MM/YYYY"),
                        "Data_Conclusao_Tarefa": st.column_config.DateColumn("Data de Conclusão", format="DD/MM/YYYY"),
                        "Status_Prazo_Tarefa": st.column_config.TextColumn("Status / Dias Restantes")
                    }
                )

            else:
                st.info("💡 **Modo de Edição Ativo:** Altere os dados abaixo e clique em 'Salvar Alterações'.")
                
                with st.form("form_edicao_projeto"):
                    st.subheader("Prazos Principais do Projeto")
                    cd1, cd2 = st.columns(2)
                    inicio_atual = info['Inicio_Projeto'].to_pydatetime() if pd.notnull(info.get('Inicio_Projeto')) else datetime.today()
                    prazo_atual = info['Prazo_Projeto'].to_pydatetime() if pd.notnull(info.get('Prazo_Projeto')) else datetime.today()
                    
                    novo_inicio_proj = cd1.date_input("Data de Início do Projeto", value=inicio_atual)
                    novo_prazo_proj = cd2.date_input("Prazo Final do Projeto", value=prazo_atual)
                    
                    st.write("---")

                    colE, colF = st.columns(2)
                    with colE:
                        novo_objetivo = st.text_area("Objetivo", value=info.get('Objetivo', ''), height=130)
                        novo_resumo = st.text_area("Resumo do Projeto", value=info.get('Resumo', ''), height=150)
                    with colF:
                        titulo_obs = info.get('Tipo_Obs', 'Observações / Justificativa')
                        nova_obs = st.text_area(f"{titulo_obs}", value=info.get('Texto_Obs', ''), height=130)
                        novos_riscos = st.text_area("Riscos", value=info.get('Riscos', ''), height=150)

                    novos_passos = st.text_area("Passos Críticos", value=info.get('Passos_Criticos', ''), height=100)
                    st.write("---")
                    
                    st.subheader("Check-list e Prazos das Tarefas")
                    df_tarefas_exibir = df_proj[df_proj['Descricao'] != ''].copy()
                    tarefas_status = {}
                    
                    h1, h2, h3 = st.columns([2, 1, 1])
                    h1.markdown("**Marque se concluído:**")
                    h2.markdown("**Alterar Prazo:**")
                    h3.markdown("**Data de Conclusão:**")
                    
                    for _, t in df_tarefas_exibir.iterrows():
                        desc = str(t['Descricao']).strip()
                        is_done = t['Concluido'] == 'Sim'
                        
                        prazo_t_atual = t['Prazo_Tarefa'].to_pydatetime() if pd.notnull(t['Prazo_Tarefa']) else datetime.today()
                        conclusao_t_atual = t['Data_Conclusao_Tarefa'].to_pydatetime() if pd.notnull(t['Data_Conclusao_Tarefa']) else None
                        
                        c1, c2, c3 = st.columns([2, 1, 1])
                        with c1:
                            done_new = st.checkbox(f"{desc}", value=is_done, help=f"Status atual: {t['Status_Prazo_Tarefa']}")
                        with c2:
                            prazo_new = st.date_input(f"Prazo", value=prazo_t_atual, key=f"dt_p_{desc}", label_visibility="collapsed")
                        with c3:
                            conclusao_new = st.date_input(f"Conclusão", value=conclusao_t_atual, key=f"dt_c_{desc}", label_visibility="collapsed")
                        
                        tarefas_status[desc] = {'concluido': done_new, 'prazo': prazo_new, 'data_conclusao': conclusao_new}

                    st.write("---")
                    st.subheader("➕ Adicionar Nova Tarefa (Opcional)")
                    col_nt1, col_nt2 = st.columns([3, 1])
                    nova_tarefa_desc = col_nt1.text_input("Descrição da Nova Tarefa")
                    nova_tarefa_prazo = col_nt2.date_input("Prazo da Nova Tarefa", value=None)

                    st.write("---")
                    col_btn1, col_btn2 = st.columns(2)
                    with col_btn1:
                        submit_salvar = st.form_submit_button("Salvar Alterações e Fechar", use_container_width=True)
                    with col_btn2:
                        submit_concluir = st.form_submit_button("Concluir Projeto Inteiro!", type="primary", use_container_width=True)

                if submit_salvar or submit_concluir:
                    with st.spinner('Salvando no Excel e Atualizando Painel...'):
                        dados_texto = {
                            'objetivo': novo_objetivo,
                            'resumo': novo_resumo,
                            'texto_obs': nova_obs,
                            'riscos': novos_riscos,
                            'passos_criticos': novos_passos,
                            'inicio_proj': novo_inicio_proj,
                            'prazo_proj': novo_prazo_proj
                        }
                        
                        sucesso, msg = salvar_alteracoes_no_excel(
                            st.session_state.projeto_ativo_state, 
                            dados_editados=dados_texto, 
                            status_tarefas=tarefas_status, 
                            concluir_projeto=submit_concluir,
                            nova_tarefa_desc=nova_tarefa_desc,
                            nova_tarefa_prazo=nova_tarefa_prazo
                        )
                        
                    if sucesso:
                        st.session_state.modo_edicao = False
                        st.success("✅ " + msg)
                        st.rerun() 
                    else:
                        st.error(f"❌ Erro: {msg}")

        st.stop()

    # =====================================================================
    # --- TELA 1: VISÃO GERAL COM ABAS E BORDA INFERIOR ---
    # =====================================================================
    st.markdown("<hr style='border: 1px solid rgba(150,150,150,0.3); margin-top: 0; margin-bottom: 1rem;'>", unsafe_allow_html=True)
    st.markdown("<h1 style='text-align: center; padding-bottom: 0;'>PAINEL DE GESTÃO CEPA</h1>", unsafe_allow_html=True)
    st.markdown("<hr style='border: 1px solid rgba(150,150,150,0.3); margin-top: 1rem; margin-bottom: 2.5rem;'>", unsafe_allow_html=True)

    def criar_card_kpi(titulo, valor):
        return f"""
        <div style="border: 1px solid rgba(150, 150, 150, 0.3); border-radius: 10px; padding: 20px 10px; text-align: center; margin-bottom: 1rem; box-shadow: 0 2px 4px rgba(0,0,0,0.05);">
            <div style="font-size: 0.9rem; font-weight: 600; opacity: 0.7; margin-bottom: 10px; text-transform: uppercase; letter-spacing: 0.5px;">{titulo}</div>
            <div style="font-size: 2.5rem; font-weight: 700; line-height: 1;">{valor}</div>
        </div>
        """

     # === SUBSTITUA ESTE BLOCO ===
    c1, c2, c3, c4 = st.columns(4)
    
    # 1. Criamos placeholders vazios para os cards.
    card_c1 = c1.empty()
    card_c2 = c2.empty()
    card_c3 = c3.empty()
    card_c4 = c4.empty()

    # 2. Inserimos o botão "Apenas Em Andamento" exatamente na coluna 3.
    with c3:
        apenas_em_andamento = st.toggle("Apenas Em Andamento", value=True)

    # 3. Lógica de Filtragem dos Dados
    df_proj_unicos = df_f.drop_duplicates(subset=['Projeto'])
    
    if apenas_em_andamento:
        # Filtra projetos E tarefas ativas (exclui projetos concluídos)
        df_atrasados = df_proj_unicos[
            (df_proj_unicos['Dias_Atraso'] > 0) & 
            (df_proj_unicos['Status_Projeto'].str.strip().str.lower() != 'concluído')
        ]
        
        # Filtra tarefas: tem que ter "Atraso", a tarefa NÃO pode estar concluída, 
        # e o projeto NÃO pode estar concluído
        tar_atrasadas = len(df_f[
            (df_f['Status_Prazo_Tarefa'].str.contains("Atraso", case=False, na=False)) & 
            (df_f['Concluido'].str.strip().str.lower() != 'sim') & 
            (df_f['Status_Projeto'].str.strip().str.lower() != 'concluído')
        ])
    else:
        # Mostra todo o histórico de atrasos (incluindo o que já foi finalizado)
        df_atrasados = df_proj_unicos[df_proj_unicos['Dias_Atraso'] > 0]
        
        # Considera todas as tarefas que carregam o status de Atraso
        tar_atrasadas = len(df_f[
            (df_f['Status_Prazo_Tarefa'].str.contains("Atraso", case=False, na=False))
        ])
        
    proj_atrasados = len(df_atrasados)

    # 4. Injetamos os cards nos espaços que reservamos lá no início
    card_c1.markdown(criar_card_kpi("PROJETOS", df_f['Projeto'].nunique()), unsafe_allow_html=True)
    card_c2.markdown(criar_card_kpi("TAREFAS", len(df_f)), unsafe_allow_html=True)
    card_c3.markdown(criar_card_kpi("PROJETOS COM ATRASO", proj_atrasados), unsafe_allow_html=True)
    card_c4.markdown(criar_card_kpi("TAREFAS PENDENTES EM ATRASO", tar_atrasadas), unsafe_allow_html=True)
    # ==============================
    st.write("")
    
    # --- Abas com persistência de estado e borda inferior (CSS já aplicado) ---
    opcoes_abas = ["Visão Geral", "Cronograma & Prazos", "Evolução & Ritmo", "Calendário", "Dados Detalhados"]
    
    if hasattr(st, "segmented_control"):
        aba_selecionada = st.segmented_control(
            "Selecione a visualização:",
            options=opcoes_abas,
            default=st.session_state.aba_ativa,
            key="aba_principal",
            label_visibility="collapsed"
        )
        if aba_selecionada:
            st.session_state.aba_ativa = aba_selecionada
    else:
        aba_selecionada = st.radio(
            "Selecione a visualização:",
            opcoes_abas,
            index=opcoes_abas.index(st.session_state.aba_ativa),
            horizontal=True,
            key="aba_principal",
            label_visibility="collapsed"
        )
        st.session_state.aba_ativa = aba_selecionada
    st.markdown("<hr style='border: 1px solid rgba(150,150,150,0.3); margin-top: -1.05rem; margin-bottom: 2rem;'>", unsafe_allow_html=True)

    # Conteúdo das abas (usando st.session_state.aba_ativa) - igual ao anterior
    if st.session_state.aba_ativa == "Visão Geral":
        co1, co2 = st.columns(2)
        with co1:
            st.subheader("Status das Tarefas")
            colors = {"Em Atraso":"#D32F2F", "Adiantado":"#388E3C", "No Prazo":"#1976D2", "Outros":"#757575"}
            def cat_st(s): 
                if "Atraso" in s: return "Em Atraso"
                elif "Adiantad" in s: return "Adiantado"
                elif "No Prazo" in s or "Restantes" in s: return "No Prazo"
                else: return "Outros"
            df_f['Cat_Prazo'] = df_f['Status_Prazo_Tarefa'].apply(cat_st)
            fig = px.pie(df_f, names='Cat_Prazo', color='Cat_Prazo', color_discrete_map=colors, hole=0.5)
            fig.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')
            st.plotly_chart(fig, use_container_width=True)
            
        with co2:
            st.subheader("Métricas e Análises") 
            metrica_selecionada = st.selectbox("Visualizar:", ["Volume de Projetos", "Progresso Médio (%)", "Esforço (Planejado vs Atraso)", "Volume de Tarefas", "Projetos por Prioridade"], key="metrica_selecionada")

            if metrica_selecionada == "Projetos por Prioridade":
                df_proj_pri = df_proj_unicos.copy()
                df_proj_pri['Classe'] = df_proj_pri['Classe'].apply(lambda x: str(x).title() if pd.notnull(x) else 'Média')
                df_pri_cnt = df_proj_pri['Classe'].value_counts().reset_index()
                df_pri_cnt.columns = ['Prioridade', 'Quantidade']
                pri_order = ['Alta', 'Média', 'Baixa']
                pri_colors = {'Alta': '#D32F2F', 'Média': '#F57C00', 'Baixa': '#757575'}
                fig = px.bar(df_pri_cnt, x='Quantidade', y='Prioridade', orientation='h', color='Prioridade', color_discrete_map=pri_colors, category_orders={'Prioridade': pri_order}, text='Quantidade')
                fig.update_traces(textposition='outside')
                fig.update_yaxes(autorange="reversed")
                fig.update_layout(showlegend=False, xaxis_title="Quantidade de Projetos", yaxis_title="Nível de Prioridade", paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')
                st.plotly_chart(fig, use_container_width=True)
            else:
                df_analise = df_f.copy()
                df_analise['Num_Área'] = df_analise['Área'].str.count(';') + 1
                df_analise['Peso_Rateio'] = 1 / df_analise['Num_Área']
                df_exp = df_analise.assign(Cli=df_analise['Área'].str.split(';')).explode('Cli')
                df_exp['Cli'] = df_exp['Cli'].str.strip()
                
                if metrica_selecionada == "Esforço (Planejado vs Atraso)":
                    df_proj_unicos_cli = df_exp.drop_duplicates(subset=['Projeto', 'Cli']).copy()
                    df_proj_unicos_cli['Plan_Pond'] = df_proj_unicos_cli['Dias_Planejados'] * df_proj_unicos_cli['Peso_Rateio']
                    df_proj_unicos_cli['Atraso_Pond'] = df_proj_unicos_cli['Dias_Atraso'] * df_proj_unicos_cli['Peso_Rateio']
                    df_chart = df_proj_unicos_cli.groupby('Cli')[['Plan_Pond', 'Atraso_Pond']].sum().reset_index()
                    df_chart['Total'] = df_chart['Plan_Pond'] + df_chart['Atraso_Pond']
                    df_chart = df_chart.sort_values('Total', ascending=True)
                    fig = go.Figure()
                    fig.add_trace(go.Bar(y=df_chart['Cli'], x=df_chart['Plan_Pond'], name='Dias Planejados', orientation='h', marker_color='#FBC02D', text=df_chart['Plan_Pond'].round(0), textposition='auto'))
                    fig.add_trace(go.Bar(y=df_chart['Cli'], x=df_chart['Atraso_Pond'], name='Dias de Atraso', orientation='h', marker_color='#D32F2F', text=df_chart['Atraso_Pond'].apply(lambda x: f"+{x:.0f}" if x > 0 else ""), textposition='inside'))
                    fig.update_layout(barmode='stack', xaxis_title="Dias Totais (Rateado)", yaxis_title="Área", legend=dict(orientation="h", y=1.1), paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')
                    st.plotly_chart(fig, use_container_width=True)
                else:
                    if metrica_selecionada == "Volume de Projetos":
                        df_chart = df_exp.drop_duplicates(subset=['Projeto', 'Cli']).groupby('Cli')['Peso_Rateio'].sum().reset_index().rename(columns={'Peso_Rateio': 'Valor'})
                        total_vol = df_chart['Valor'].sum()
                        if total_vol > 0: df_chart['Valor'] = (df_chart['Valor'] / total_vol) * 100
                        text_template, eixo_x_title, cor_barra = '%{value:.1f}%', "Proporção do Volume de Projetos (%)", '#1976D2' 
                    elif metrica_selecionada == "Progresso Médio (%)":
                        df_chart = df_exp.drop_duplicates(subset=['Projeto', 'Cli']).groupby('Cli')['Percentual_Conclusao'].mean().reset_index()
                        df_chart['Valor'] = (df_chart['Percentual_Conclusao'] * 100).round(1)
                        text_template, eixo_x_title, cor_barra = '%{value}%', "Média de Conclusão (%)", '#388E3C' 
                    elif metrica_selecionada == "Volume de Tarefas":
                        df_chart = df_exp.groupby('Cli')['Peso_Rateio'].sum().reset_index().rename(columns={'Peso_Rateio': 'Valor'})
                        total_vol = df_chart['Valor'].sum()
                        if total_vol > 0: df_chart['Valor'] = (df_chart['Valor'] / total_vol) * 100
                        text_template, eixo_x_title, cor_barra = '%{value:.1f}%', "Proporção do Volume de Tarefas (%)", '#7B1FA2' 

                    df_chart = df_chart.sort_values('Valor', ascending=True)
                    fig = px.bar(df_chart, x='Valor', y='Cli', orientation='h', text='Valor')
                    fig.update_traces(marker_color=cor_barra, texttemplate=text_template, textposition='outside')
                    fig.update_layout(xaxis_title=eixo_x_title, yaxis_title="Área", paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')
                    st.plotly_chart(fig, use_container_width=True)

    elif st.session_state.aba_ativa == "Cronograma & Prazos":
        df_gantt = df_f.drop_duplicates(subset=['Projeto']).copy()
        
        col_t2_1, col_t2_2, col_t2_3 = st.columns([2, 1, 1])
        with col_t2_1: st.subheader("Linha do Tempo de Projetos")
        
        status_opcoes = sorted(df_gantt['Status_Projeto'].unique().tolist())
        with col_t2_2: 
            sts_gantt = st.multiselect("Filtrar Status Específico:", status_opcoes, default=[], placeholder="Todos os Status", key="filtro_sts_gantt")
        with col_t2_3:
            ordem = st.selectbox("Ordenar por:", ["Data Início (Mais Cedo)", "Prioridade (Alta > Baixa)", "Data Início (Mais Tarde)", "Prazo Final", "Maior Progresso", "Maior Atraso"], key="gantt_sort")
            apenas_atrasados = st.checkbox("Apenas c/ Atraso", key="filtro_atraso_gantt")
        
        if sts_gantt: df_gantt = df_gantt[df_gantt['Status_Projeto'].isin(sts_gantt)]
        if apenas_atrasados: df_gantt = df_gantt[df_gantt['Dias_Atraso'] > 0]

        if not df_gantt.empty and 'Inicio_Projeto' in df_gantt.columns and 'Prazo_Projeto' in df_gantt.columns:
            hoje = datetime.today()
            df_gantt['Inicio_Projeto'] = df_gantt['Inicio_Projeto'].fillna(hoje)
            df_gantt['Prazo_Projeto'] = df_gantt['Prazo_Projeto'].fillna(hoje)
            df_gantt['Inicio_Atraso'] = df_gantt['Prazo_Projeto']
            def get_fim_atraso(row):
                if row['Dias_Atraso'] > 0: return row['Data_Conclusao_Projeto'] if pd.notnull(row['Data_Conclusao_Projeto']) else datetime.today()
                return row['Prazo_Projeto']
            df_gantt['Fim_Atraso_Visual'] = df_gantt.apply(get_fim_atraso, axis=1)
            df_gantt['Fim_Total_Real'] = df_gantt[['Prazo_Projeto', 'Fim_Atraso_Visual']].max(axis=1)
            df_gantt['Duracao_Total_Dias'] = (df_gantt['Fim_Total_Real'] - df_gantt['Inicio_Projeto']).dt.days.replace(0, 1)
            df_gantt['Fim_Progresso'] = df_gantt['Inicio_Projeto'] + pd.to_timedelta(df_gantt['Duracao_Total_Dias'] * df_gantt['Percentual_Conclusao'], unit='D')
            
            def format_gantt_y_axis(row):
                classe_str = str(row.get('Classe', 'Média')).title()
                inicial = classe_str[0] if classe_str else 'M'
                cor_txt = "#D32F2F" if classe_str == 'Alta' else ("#757575" if classe_str == 'Baixa' else "#F57C00")
                nome_proj = row['Projeto']
                link = urllib.parse.quote(nome_proj)
                return f"<a href='?projeto={link}' target='_self' style='color:{cor_txt}; text-decoration:none;'><b>[{inicial}]</b> {nome_proj}</a>"
            
            df_gantt['Projeto_Display'] = df_gantt.apply(format_gantt_y_axis, axis=1)
            map_prioridade = {'Alta': 3, 'Média': 2, 'Baixa': 1}
            df_gantt['Pri_Num'] = df_gantt['Classe'].apply(lambda x: map_prioridade.get(str(x).title(), 2))
            
            if ordem == "Prioridade (Alta > Baixa)": df_gantt = df_gantt.sort_values(by=['Pri_Num', 'Inicio_Projeto'], ascending=[True, False])
            elif ordem == "Data Início (Mais Cedo)": df_gantt = df_gantt.sort_values('Inicio_Projeto', ascending=True)
            elif ordem == "Data Início (Mais Tarde)": df_gantt = df_gantt.sort_values('Inicio_Projeto', ascending=False)
            elif ordem == "Prazo Final": df_gantt = df_gantt.sort_values('Prazo_Projeto', ascending=True)
            elif ordem == "Maior Progresso": df_gantt = df_gantt.sort_values('Percentual_Conclusao', ascending=False)
            elif ordem == "Maior Atraso": df_gantt = df_gantt.sort_values('Dias_Atraso', ascending=False)

            fig_total = px.timeline(df_gantt, x_start="Inicio_Projeto", x_end="Prazo_Projeto", y="Projeto_Display", hover_data={"Percentual_Conclusao":':.0%'})
            fig_total.update_traces(marker_color='#757575', marker_line_width=0, opacity=0.6, name="Prazo Planejado")

            df_atrasados = df_gantt[df_gantt['Dias_Atraso'] > 0].copy()
            fig_atraso = px.timeline(df_atrasados, x_start="Inicio_Atraso", x_end="Fim_Atraso_Visual", y="Projeto_Display", hover_data={"Dias_Atraso":':.0f'})
            fig_atraso.update_traces(marker_color='#D32F2F', opacity=0.8, name="Atraso")

            df_progresso = df_gantt[df_gantt['Percentual_Conclusao'] > 0.01].copy()
            colors_map = {"Concluído": "#388E3C", "Em Andamento": "#00A9CC", "Atrasado": "#D32F2F", "Não Iniciado": "#424242"}
            fig_prog = px.timeline(df_progresso, x_start="Inicio_Projeto", x_end="Fim_Progresso", y="Projeto_Display", color="Status_Projeto", color_discrete_map=colors_map, hover_data={"Percentual_Conclusao":':.0%'})
            
            fig = go.Figure(data=fig_total.data + fig_atraso.data + fig_prog.data)
            fig.update_layout(barmode='overlay', height=max(400, len(df_gantt) * 40), xaxis_title="Calendário", yaxis_title="", legend=dict(orientation="h", y=1.1), xaxis=dict(type='date', tickformat="%b/%Y"), paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')
            fig.update_yaxes(autorange="reversed") 
            st.plotly_chart(fig, use_container_width=True)
            st.caption("Eixo Y: Cores das letras refletem Prioridade. **Clique no nome do Projeto para abrir a ficha de edição.**")

    elif st.session_state.aba_ativa == "Evolução & Ritmo":
        col_ritmo_1, col_ritmo_2 = st.columns([1, 1])
        with col_ritmo_1: freq_escolhida = st.selectbox("Periodicidade:", ["Semanal", "Quinzenal", "Mensal"], key="freq_ritmo")
        with col_ritmo_2: tipo_dado = st.selectbox("Analisar:", ["Projetos", "Tarefas"], key="tipo_ritmo")
        
        if tipo_dado == "Projetos":
            df_ritmo = df_f.drop_duplicates(subset=['Projeto']).copy()
            col_prazo, col_real = 'Prazo_Projeto', 'Data_Conclusao_Projeto'
        else:
            df_ritmo = df_f.copy()
            col_prazo, col_real = 'Prazo_Tarefa', 'Data_Conclusao_Tarefa'

        freq_pd = 'W-MON' if freq_escolhida == "Semanal" else ('2W-MON' if freq_escolhida == "Quinzenal" else 'ME')
        lbl_format = '%b/%Y'

        df_plan = df_ritmo[df_ritmo[col_prazo].notnull()].copy()
        serie_plan = df_plan.groupby(pd.Grouper(key=col_prazo, freq=freq_pd)).size()
        df_real = df_ritmo[df_ritmo[col_real].notnull()].copy()
        serie_real = df_real.groupby(pd.Grouper(key=col_real, freq=freq_pd)).size()

        df_chart_ritmo = pd.DataFrame({'Planejado': serie_plan, 'Realizado': serie_real}).fillna(0).sort_index().cumsum()
        df_real_viz = df_chart_ritmo[df_chart_ritmo.index <= pd.Timestamp.today().normalize()]

        fig_ritmo = go.Figure()
        fig_ritmo.add_trace(go.Scatter(x=df_chart_ritmo.index, y=df_chart_ritmo['Planejado'], mode='lines+markers', name=f'{tipo_dado} Planejados', line=dict(color='#1976D2', width=3), marker=dict(size=8)))
        fig_ritmo.add_trace(go.Scatter(x=df_real_viz.index, y=df_real_viz['Realizado'], mode='lines+markers', name=f'{tipo_dado} Entregues', line=dict(color='#388E3C', width=3), fill='tozeroy', fillcolor='rgba(56, 142, 60, 0.2)', marker=dict(size=8)))
        fig_ritmo.update_layout(title=f"Curva S: {tipo_dado} ({freq_escolhida})", xaxis_title="Tempo", yaxis_title="Acumulado", hovermode="x unified", xaxis=dict(tickformat=lbl_format, tickmode='auto'), legend=dict(orientation="h", y=1.1), paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')
        st.plotly_chart(fig_ritmo, use_container_width=True)

    elif st.session_state.aba_ativa == "Calendário":
        col_cal, col_crit = st.columns([3, 1])

        with col_cal:
            st.subheader("Calendário de Entregas")
            c1_cal, c2_cal = st.columns(2)
            hoje_cal = datetime.today()
            
            meses = {1:"Janeiro", 2:"Fevereiro", 3:"Março", 4:"Abril", 5:"Maio", 6:"Junho", 7:"Julho", 8:"Agosto", 9:"Setembro", 10:"Outubro", 11:"Novembro", 12:"Dezembro"}
            mes_sel = c1_cal.selectbox("Mês", options=list(meses.keys()), format_func=lambda x: meses[x], index=hoje_cal.month-1)
            ano_sel = c2_cal.selectbox("Ano", options=range(hoje_cal.year-2, hoje_cal.year+3), index=2)
            
            st.markdown("<div style='margin-top: 10px;'></div>", unsafe_allow_html=True)
            esconder_concluidas = st.checkbox("Ocultar tarefas concluídas")

            df_mes = df_f[df_f['Prazo_Tarefa'].dt.month == mes_sel].copy()
            df_mes = df_mes[df_mes['Prazo_Tarefa'].dt.year == ano_sel]
            
            if esconder_concluidas:
                df_mes = df_mes[df_mes['Concluido'].str.lower() != 'sim']

            html = "<table style='width: 100%; border-collapse: collapse; table-layout: fixed; margin-top: 15px;'><tr>"
            for d in ["Seg", "Ter", "Qua", "Qui", "Sex", "Sáb", "Dom"]: html += f"<th style='border: 1px solid gray; padding: 8px; text-align: center;'>{d}</th>"
            html += "</tr><tr>"

            primeiro_dia, num_dias = calendar.monthrange(ano_sel, mes_sel)

            for _ in range(primeiro_dia): html += "<td style='border: 1px solid gray; padding: 8px; height: 110px;'></td>"

            dia_atual = 1
            col_idx = primeiro_dia

            while dia_atual <= num_dias:
                tarefas_dia = df_mes[df_mes['Prazo_Tarefa'].dt.day == dia_atual]
                tarefas_html = ""

                for _, t in tarefas_dia.iterrows():
                    concluido = str(t.get('Concluido', 'Não')).strip().lower()
                    status_prazo = str(t['Status_Prazo_Tarefa']).lower()
                    cor_bg = "#388E3C" if concluido == 'sim' else ("#D32F2F" if "atraso" in status_prazo else "#1976D2")
                    
                    desc_tarefa = str(t['Descricao'])
                    nome_proj = str(t['Projeto'])
                    
                    classe_str = str(t.get('Classe', 'Média')).title()
                    inicial_classe = classe_str[0] if classe_str else 'M'
                    cor_pri = "#B71C1C" if classe_str == 'Alta' else ("#455A64" if classe_str == 'Baixa' else "#F57C00")
                        
                    badge_html = f"<span style='background-color:{cor_pri}; color:white; padding: 0px 4px; margin-right: 4px; border-radius: 2px; font-size: 9px; font-weight: bold; border: 1px solid rgba(255,255,255,0.7); display: inline-block; vertical-align: middle;'>{inicial_classe}</span>"
                    hover_text = f"Clique para Detalhes | Prioridade: {classe_str} | {nome_proj} - {desc_tarefa}"
                    
                    link_seguro = urllib.parse.quote(nome_proj)
                    div_tarefa = f"<div style='background-color: {cor_bg}; color: white; border-radius: 4px; padding: 3px 5px; margin-bottom: 3px; font-size: 11px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; cursor: pointer;' title='{hover_text}'>{badge_html}<b>{nome_proj[:8]}</b>: {desc_tarefa}</div>"
                    tarefas_html += f"<a href='?projeto={link_seguro}' target='_self' style='text-decoration: none;'>{div_tarefa}</a>"

                is_today = (dia_atual == hoje_cal.day and mes_sel == hoje_cal.month and ano_sel == hoje_cal.year)
                bg_td = "rgba(25, 118, 210, 0.2)" if is_today else "transparent"
                peso_fonte = "bold" if is_today else "normal"

                html += f"<td style='border: 1px solid gray; padding: 4px; height: 110px; vertical-align: top; background-color: {bg_td};'><div style='font-weight: {peso_fonte}; margin-bottom: 4px; text-align: right;'>{dia_atual}</div>{tarefas_html}</td>"

                dia_atual += 1
                col_idx += 1
                if col_idx == 7 and dia_atual <= num_dias:
                    html += "</tr><tr>"
                    col_idx = 0

            while col_idx < 7 and col_idx > 0:
                html += "<td style='border: 1px solid gray; padding: 8px; height: 110px;'></td>"
                col_idx += 1

            html += "</tr></table>"
            st.markdown(html, unsafe_allow_html=True)
            st.caption("Fundo: [Verde] Concluído | [Azul] Planejado | [Vermelho] Atrasado. *Clique na tarefa para ver a ficha completa do projeto.*")

        with col_crit:
            st.subheader("Tarefas Críticas")
            filtro_crit = st.radio("Visualizar:", ["Vencendo em até 7 dias", "Apenas Atrasadas", "Todas Pendentes"])
            
            df_crit = df_f[df_f['Concluido'].astype(str).str.strip().str.lower() != 'sim'].copy()
            df_crit = df_crit[df_crit['Prazo_Tarefa'].notnull()]
            df_crit['Dias_Restantes'] = (df_crit['Prazo_Tarefa'] - hoje_cal).dt.days
            
            if filtro_crit == "Vencendo em até 7 dias": df_crit = df_crit[(df_crit['Dias_Restantes'] >= 0) & (df_crit['Dias_Restantes'] <= 7)]
            elif filtro_crit == "Apenas Atrasadas": df_crit = df_crit[df_crit['Dias_Restantes'] < 0]
            
            df_crit = df_crit.sort_values('Dias_Restantes', ascending=True)

            with st.container(height=650):
                if df_crit.empty:
                    st.success("Nenhuma tarefa encontrada.")
                else:
                    for _, t in df_crit.iterrows():
                        dias = t['Dias_Restantes']
                        nome_proj = t['Projeto']
                        desc = t['Descricao']
                        data_formatada = t['Prazo_Tarefa'].strftime('%d/%m/%Y')
                        classe_proj = str(t.get('Classe', 'Média')).title()
                        
                        if dias < 0: cor_hex, tag = "#D32F2F", f"[ATRASADO {-dias} DIAS]"
                        elif dias == 0: cor_hex, tag = "#F57C00", "[VENCE HOJE]"
                        elif dias <= 7: cor_hex, tag = "#FBC02D", f"[VENCE EM {dias} DIAS]"
                        else: cor_hex, tag = "#388E3C", f"[VENCE EM {dias} DIAS]"

                        st.markdown(f"<span style='color:{cor_hex}; font-size:1.2em;'>●</span> **{tag}** {desc}", unsafe_allow_html=True)
                        st.caption(f"{nome_proj} | {data_formatada}")
                        st.divider()

    elif st.session_state.aba_ativa == "Dados Detalhados":
        cols_show = ['Projeto', 'Área', 'Classe', 'Inicio_Projeto', 'Prazo_Projeto', 'Dias_Planejados', 'Dias_Atraso', 'Percentual_Conclusao', 'Status_Projeto']
        final_cols = [c for c in cols_show if c in df_f.columns]
        df_show = df_f[final_cols].copy()
        
        if 'Percentual_Conclusao' in df_show.columns: 
            df_show['Percentual_Conclusao'] = df_show['Percentual_Conclusao'] * 100
            
        st.dataframe(
            df_show,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Inicio_Projeto": st.column_config.DateColumn("Início", format="DD/MM/YYYY"),
                "Prazo_Projeto": st.column_config.DateColumn("Prazo Final", format="DD/MM/YYYY"),
                "Dias_Planejados": st.column_config.NumberColumn("Dias Planejados", format="%d"),
                "Dias_Atraso": st.column_config.NumberColumn("Atraso (Dias)", format="%d"),
                "Percentual_Conclusao": st.column_config.ProgressColumn("Progresso", format="%d%%", min_value=0, max_value=100),
                "Projeto": st.column_config.TextColumn("Projeto"),
                "Área": st.column_config.TextColumn("Área(s)"),
                "Classe": st.column_config.TextColumn("Prioridade"),
                "Status_Projeto": st.column_config.TextColumn("Status")
            }
        )

else:
    st.warning("Arquivo 'Base_Dados_Consolidada_CEPA.xlsx' não encontrado na pasta do aplicativo.")