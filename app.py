import streamlit as st
import pandas as pd
from simple_salesforce import Salesforce
from io import BytesIO
from datetime import datetime, timedelta, timezone
import plotly.express as px
import time

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Gestão de Casos", layout="wide", initial_sidebar_state="expanded")

# --- CSS CUSTOMIZADO ---
st.markdown("""
    <style>
    .stApp { background-color: #ffffff !important; }
    
    /* CORREÇÃO DA BARRA LATERAL: Fundo transparente em vez de invisível */
    header { background-color: transparent !important; }
    
    #MainMenu { visibility: hidden !important; display: none !important; }
    footer { visibility: hidden !important; display: none !important; }
    .block-container { padding-top: 1.5rem !important; padding-bottom: 1rem !important; }
    
    h1 { font-size: 24px !important; margin-bottom: 20px !important; color: #1a2935; }
    
    .stButton>button {
        border-radius: 0px 0px 8px 8px !important;
        border-top: none !important;
        background-color: #f1f3f5 !important;
        color: #0056b3 !important;
        font-weight: 600 !important;
        margin-top: -15px !important;
        border: 1px solid #dce1e6 !important;
    }
    .stButton>button:hover {
        background-color: #0056b3 !important;
        color: white !important;
    }
    .btn-login>button {
        border-radius: 8px !important;
        margin-top: 10px !important;
        background-color: #0056b3 !important;
        color: white !important;
        border: none !important;
    }
    .btn-login>button:hover {
        background-color: #004494 !important;
    }
    </style>
""", unsafe_allow_html=True)

# --- FUSO HORÁRIO BRASIL (UTC -3) ---
fuso_br = timezone(timedelta(hours=-3))

# --- CONTROLE DE ESTADO & SESSÃO ---
if 'fila_selecionada' not in st.session_state:
    st.session_state.fila_selecionada = None
if 'last_update' not in st.session_state:
    st.session_state.last_update = datetime.now(fuso_br).strftime("%d/%m/%Y %H:%M")
if 'sf_authenticated' not in st.session_state:
    st.session_state.sf_authenticated = False
if 'sf_username' not in st.session_state:
    st.session_state.sf_username = ""
if 'sf_password' not in st.session_state:
    st.session_state.sf_password = ""
if 'sf_token' not in st.session_state:
    st.session_state.sf_token = ""

# --- TELA DE LOGIN ---
if not st.session_state.sf_authenticated:
    col_vazia1, col_login, col_vazia2 = st.columns([1, 2, 1])
    with col_login:
        st.markdown("<h2 style='text-align: center; color: #0c1c2b; margin-top: 50px;'>🔐 Login Operacional</h2>", unsafe_allow_html=True)
        st.markdown("<p style='text-align: center; color: #6c757d; margin-bottom: 30px;'>Conecte-se com suas credenciais do Salesforce para gerenciar os casos em seu nome.</p>", unsafe_allow_html=True)
        
        with st.form("login_form"):
            st.info("ℹ️ **Segurança:** O painel NÃO salva sua senha em nenhum banco de dados. As credenciais são usadas apenas enquanto essa janela estiver aberta.")
            user_input = st.text_input("👤 Usuário (E-mail Salesforce)")
            pwd_input = st.text_input("🔑 Senha", type="password")
            token_input = st.text_input("🛡️ Token de Segurança", type="password")
            
            st.markdown("<div class='btn-login'>", unsafe_allow_html=True)
            submitted = st.form_submit_button("Entrar no Painel", use_container_width=True)
            st.markdown("</div>", unsafe_allow_html=True)
            
            if submitted:
                if user_input and pwd_input and token_input:
                    try:
                        with st.spinner("Autenticando no Salesforce..."):
                            Salesforce(username=user_input, password=pwd_input, security_token=token_input, domain='login')
                            st.session_state.sf_authenticated = True
                            st.session_state.sf_username = user_input
                            st.session_state.sf_password = pwd_input
                            st.session_state.sf_token = token_input
                            st.rerun()
                    except Exception as e:
                        st.error("❌ Falha na autenticação. Verifique seu usuário, senha e token de segurança.")
                else:
                    st.warning("⚠️ Por favor, preencha todos os campos.")
    st.stop()

# --- CONEXÃO DINÂMICA ---
@st.cache_resource
def init_connection(user, pwd, token):
    return Salesforce(username=user, password=pwd, security_token=token, domain='login')

sf = init_connection(st.session_state.sf_username, st.session_state.sf_password, st.session_state.sf_token)

# --- FUNÇÕES DE BUSCA ---
@st.cache_data(ttl=86400)
def get_api_user_id(username, _pwd, _token):
    sf_conn = init_connection(username, _pwd, _token)
    try:
        res = sf_conn.query(f"SELECT Id FROM User WHERE Username = '{username}'")
        if res['totalSize'] > 0:
            return res['records'][0]['Id']
    except Exception as e:
        pass
    return None

@st.cache_data(ttl=3600, show_spinner="Atualizando lista de proprietários...") 
def get_owner_options(username, _pwd, _token):
    sf_conn = init_connection(username, _pwd, _token)
    users = sf_conn.query_all("SELECT Id, Name FROM User WHERE IsActive = TRUE")
    queues = sf_conn.query_all("SELECT Id, Name FROM Group WHERE Type = 'Queue'")
    
    opcoes = {}
    for q in queues['records']:
        opcoes[f"📁 {q['Name']}"] = q['Id']
    for u in users['records']:
        opcoes[f"👤 {u['Name']}"] = u['Id']
        
    return dict(sorted(opcoes.items()))

@st.cache_data(ttl=1800, show_spinner=False) 
def get_data(periodo_selecionado, dt_inicio, dt_fim, incluir_fechados, username, _pwd, _token):
    sf_conn = init_connection(username, _pwd, _token)
    
    if periodo_selecionado == "Personalizado" and dt_inicio and dt_fim:
        inicio_str = dt_inicio.strftime('%Y-%m-%dT00:00:00Z')
        fim_str = dt_fim.strftime('%Y-%m-%dT23:59:59Z')
        filtro_data = f"CreatedDate >= {inicio_str} AND CreatedDate <= {fim_str}"
    else:
        mapa_periodos = {
            "Últimos 30 Dias": "LAST_N_DAYS:30",
            "Últimos 60 Dias": "LAST_N_DAYS:60",
            "Últimos 90 Dias": "LAST_N_DAYS:90",
            "Este Ano": "THIS_YEAR"
        }
        filtro_data = f"CreatedDate = {mapa_periodos.get(periodo_selecionado, 'LAST_N_DAYS:30')}"

    filtro_status = "" if incluir_fechados else "AND Status != 'Closed' AND Status != 'Fechado'"

    query = f"""
    SELECT 
        Id, CaseNumber, CreatedDate, ClosedDate, Status, Description,
        Account.Name, Account.FOZ_CPF__c, Account.FOZ_Classificacao__c,
        Origin, Type, FOZ_TipoSolicitacao__c, FOZ_Motivo__c, FOZ_Detalhe__c, FOZ_Subdetalhe__c, OwnerId, Owner.Name, 
        (SELECT IsViolated FROM CaseMilestones),
        (SELECT CommentBody, CreatedBy.Name, CreatedDate FROM CaseComments ORDER BY CreatedDate ASC)
    FROM Case 
    WHERE (Type = 'OA' 
           OR (Owner.Name LIKE 'CARTEIRA%' AND (Type != 'OS' OR Type = null)) 
           OR Owner.Name LIKE '%GENÉRICO%'
           OR Owner.Name LIKE '%GENERICO%'
           OR Owner.Name LIKE '%Casos sem fila%')
      AND {filtro_data}
      {filtro_status}
    """
    
    # CORREÇÃO AQUI: URL Base voltou pro código!
    sf_base_url = "https://ibbl.lightning.force.com/lightning/r/Case/"
    
    # --- INÍCIO DA BARRA DE PROGRESSO ---
    my_bar = st.progress(0, text="Iniciando sincronização com o Salesforce...")
    
    try:
        result = sf_conn.query(query)
        total_records = result.get('totalSize', 0)
        records = result.get('records', [])
        
        if total_records > 0:
            current_len = len(records)
            percent = int((current_len / total_records) * 40)
            my_bar.progress(percent, text=f"Baixando casos... ({current_len} de {total_records})")

            while not result.get('done', True):
                result = sf_conn.query_more(result['nextRecordsUrl'], True)
                records.extend(result.get('records', []))
                current_len = len(records)
                percent = int((current_len / total_records) * 40)
                my_bar.progress(percent, text=f"Baixando casos... ({current_len} de {total_records})")

        my_bar.progress(40, text="Download concluído! Estruturando inteligência de dados...")

        linhas = []
        hoje_utc = datetime.now(timezone.utc).replace(tzinfo=None)
        total_processar = len(records)
        
        for i, record in enumerate(records):
            if total_processar > 0 and i % 500 == 0:
                progresso_atual = 40 + int((i / total_processar) * 60)
                my_bar.progress(progresso_atual, text=f"Estruturando dados visuais... {progresso_atual}%")

            dono_upper = record['Owner']['Name'].upper() if record['Owner'] else 'SISTEMA/SEM DONO'
            
            filas_conhecidas = [
                "ERRO SISTÊMICO", "CAPACIDADE", "FRANQUIAS", "AUDITORIA", 
                "HELP TEC", "JURÍDICO", "INFORMAÇÃO", "RAF", "FINANCEIRO", "BACKOFFICE"
            ]
            
            if "SAFETY" in dono_upper:
                fila_principal = "SAFETY"
                subfila = dono_upper
            elif "GENÉRICO" in dono_upper or "GENERICO" in dono_upper or "SEM FILA" in dono_upper:
                fila_principal = "CASOS SEM FILA - GENÉRICO"
                subfila = dono_upper
            elif dono_upper in filas_conhecidas:
                fila_principal = dono_upper
                subfila = "-"
            elif dono_upper.startswith("CARTEIRA"):
                fila_principal = "CORPORATIVO"
                subfila = dono_upper
            else:
                fila_principal = "ATRIBUÍDO AO USUÁRIO"
                subfila = dono_upper
                
            macro_status = "🟢 Fechado" if record['Status'] in ['Closed', 'Fechado'] else "🟡 Em Tratativa"
            sla_atrasado = any(m['IsViolated'] for m in record['CaseMilestones']['records']) if record['CaseMilestones'] else False
            
            data_abertura = pd.to_datetime(record['CreatedDate']).tz_localize(None)
            data_fechamento = pd.to_datetime(record['ClosedDate']).tz_localize(None) if record.get('ClosedDate') else None
            
            fim_calc = data_fechamento if data_fechamento else hoje_utc
            idade_dias = (fim_calc - data_abertura).days
            
            classificacao = '-'
            if record['Account'] and record['Account'].get('FOZ_Classificacao__c'):
                classificacao = record['Account']['FOZ_Classificacao__c']
                
            desc_oficial = record.get('Description')
            historico_comentarios = ""
            
            if record.get('CaseComments') and record['CaseComments'].get('records'):
                for comment in record['CaseComments']['records']:
                    autor = comment['CreatedBy']['Name'] if comment.get('CreatedBy') else 'Usuário'
                    texto = comment['CommentBody']
                    try:
                        dt_obj = pd.to_datetime(comment['CreatedDate']).tz_convert(fuso_br)
                        data_str = dt_obj.strftime('%d/%m/%Y %H:%M')
                    except:
                        data_str = comment['CreatedDate']
                        
                    historico_comentarios += f"🗣️ {autor} em {data_str}:\n{texto}\n\n"
            
            if desc_oficial and historico_comentarios:
                descricao_final = f"📝 DESCRIÇÃO ORIGINAL:\n{desc_oficial}\n\n{'-'*40}\n\n💬 HISTÓRICO DE COMENTÁRIOS:\n{historico_comentarios}".strip()
            elif historico_comentarios:
                descricao_final = f"💬 HISTÓRICO DE COMENTÁRIOS:\n{historico_comentarios}".strip()
            else:
                descricao_final = desc_oficial if desc_oficial else "-"
                
            linhas.append({
                'ID do Caso': record['Id'],
                'ID do Proprietário': record['OwnerId'],
                'Link Salesforce': f"{sf_base_url}{record['Id']}/view",
                'Número': record['CaseNumber'],
                'Status': record['Status'],
                'Motivo': record['FOZ_Motivo__c'],
                'Descrição': descricao_final,
                'Fila Principal': fila_principal,
                'Subfila': subfila,
                'Macro Status': macro_status,
                'Idade (Dias)': idade_dias,
                'SLA Atrasado': '🔴 Atrasado' if sla_atrasado else '✅ No Prazo',
                'Origem': record['Origin'],
                'Tipo Solicitação': record['FOZ_TipoSolicitacao__c'],
                'Conta': record['Account']['Name'] if record['Account'] else '-',
                'Abertura': data_abertura,
                'Fechamento': data_fechamento
            })
            
        df_final = pd.DataFrame(linhas)
        
        if not df_final.empty:
            df_final['Abertura'] = pd.to_datetime(df_final['Abertura'])
            df_final['Fechamento'] = pd.to_datetime(df_final['Fechamento'])
        
        my_bar.progress(100, text="✅ Sincronização e Processamento finalizados!")
        time.sleep(0.5)
        my_bar.empty()
        
        return df_final

    except Exception as e:
        my_bar.empty()
        st.error(f"Erro de comunicação com o Salesforce: {e}")
        return pd.DataFrame()

# --- FUNÇÃO PARA DESENHAR O CARD QUADRADO ---
def desenhar_card(fila_nome, df_fila):
    vol = len(df_fila)
    trat = len(df_fila[df_fila['Macro Status'] == '🟡 Em Tratativa'])
    fech = len(df_fila[df_fila['Macro Status'] == '🟢 Fechado'])
    atr = len(df_fila[(df_fila['SLA Atrasado'] == '🔴 Atrasado') & (df_fila['Macro Status'] == '🟡 Em Tratativa')])
    
    cor_atraso = "#d9534f" if atr > 0 else "#555555"
    borda_destaque = "border-top: 4px solid #d9534f;" if fila_nome == "SAFETY" else ""
    
    html_card = f"""
    <div style="background-color: white; border: 1px solid #dce1e6; border-radius: 8px 8px 0px 0px; padding: 15px; height: 145px; box-shadow: 0 2px 4px rgba(0,0,0,0.02); display: flex; flex-direction: column; justify-content: space-between; {borda_destaque}">
        <h4 style="margin: 0; padding: 0; color: #0c1c2b; font-size: 15px; text-align: center; text-transform: uppercase; letter-spacing: 0.5px;">{fila_nome}</h4>
        <div style="font-size: 13px; color: #495057; line-height: 1.6; margin-top: 10px;">
            <div style="display: flex; justify-content: space-between;"><span>Casos:</span> <b>{vol}</b></div>
            <div style="display: flex; justify-content: space-between;"><span>Em tratativa:</span> <b>{trat}</b></div>
            <div style="display: flex; justify-content: space-between;"><span>Fechados:</span> <b>{fech}</b></div>
            <div style="display: flex; justify-content: space-between; border-top: 1px dashed #eee; margin-top: 4px; padding-top: 4px;"><span>SLA Atrasado:</span> <b style="color: {cor_atraso};">{atr}</b></div>
        </div>
    </div>
    """
    st.markdown(html_card, unsafe_allow_html=True)
    if st.button(f"Abrir Detalhe", key=f"btn_{fila_nome}", use_container_width=True):
        st.session_state.fila_selecionada = fila_nome
        st.rerun()

# --- MENU LATERAL (SIDEBAR) ---
try:
    st.sidebar.image("Salesforce.png", use_container_width=True)
except Exception:
    st.sidebar.markdown("<h2>Filtros</h2>", unsafe_allow_html=True)

st.sidebar.markdown(f"**Logado como:**<br> <span style='color: #0056b3; font-size: 14px;'>{st.session_state.sf_username}</span>", unsafe_allow_html=True)
st.sidebar.caption(f"Última Sincronização: {st.session_state.last_update}")

# Busca Global
busca_global = st.sidebar.text_input("🔍 Busca Rápida (Nº do Caso ou Conta)")

st.markdown("""
    <style>
    [data-testid="stSidebar"] div.stButton > button { width: 95% !important; border-radius: 6px !important; margin-top: 10px !important; border: none !important; }
    </style>
""", unsafe_allow_html=True)

if st.sidebar.button("🔄 Sincronizar Agora", type="primary", use_container_width=True):
    st.cache_data.clear()
    st.session_state.last_update = datetime.now(fuso_br).strftime("%d/%m/%Y %H:%M")
    st.rerun()

if st.sidebar.button("🚪 Sair (Logout)", use_container_width=True):
    st.session_state.sf_authenticated = False
    st.session_state.sf_username = ""
    st.session_state.sf_password = ""
    st.session_state.sf_token = ""
    st.session_state.fila_selecionada = None
    st.cache_data.clear()
    st.cache_resource.clear()
    st.rerun()

st.sidebar.markdown("---")
opcoes_periodo = ["Últimos 30 Dias", "Últimos 60 Dias", "Últimos 90 Dias", "Este Ano", "Personalizado"]
periodo_selecionado = st.sidebar.selectbox("Período de Abertura", opcoes_periodo, index=0)

dt_inicio, dt_fim = None, None
if periodo_selecionado == "Personalizado":
    datas = st.sidebar.date_input("Selecione o intervalo (Início e Fim)", [])
    if len(datas) == 2:
        dt_inicio, dt_fim = datas
    else:
        st.sidebar.warning("Selecione a data de início e fim no calendário para carregar.")
        st.stop()

incluir_fechados = st.sidebar.checkbox("Mostrar Casos Fechados", value=False)

user, pwd, token = st.session_state.sf_username, st.session_state.sf_password, st.session_state.sf_token
df_filtrado = get_data(periodo_selecionado, dt_inicio, dt_fim, incluir_fechados, user, pwd, token)
lista_proprietarios = get_owner_options(user, pwd, token)
api_user_id = get_api_user_id(user, pwd, token)

# Aplicação da Busca Global
if busca_global and not df_filtrado.empty:
    mask = (
        df_filtrado['Número'].astype(str).str.contains(busca_global, case=False, na=False) |
        df_filtrado['Conta'].astype(str).str.contains(busca_global, case=False, na=False)
    )
    df_filtrado = df_filtrado[mask]

# --- ORDENAÇÃO DAS FILAS ---
todas_filas = df_filtrado['Fila Principal'].unique().tolist() if not df_filtrado.empty else []
if "ATRIBUÍDO AO USUÁRIO" in todas_filas:
    todas_filas.remove("ATRIBUÍDO AO USUÁRIO")
    filas_ordenadas = sorted(todas_filas)
    filas_ordenadas.append("ATRIBUÍDO AO USUÁRIO")
else:
    filas_ordenadas = sorted(todas_filas)

# --- RENDERIZAÇÃO DA TELA PRINCIPAL ---
if df_filtrado.empty:
    st.markdown("<h1>Visão Operacional de Casos</h1>", unsafe_allow_html=True)
    st.info("Nenhum caso encontrado para os filtros selecionados.")
    
elif st.session_state.fila_selecionada is None:
    st.markdown("<h1>Visão Operacional de Casos</h1>", unsafe_allow_html=True)
    
    cols = st.columns(4)
    for i, fila in enumerate(filas_ordenadas):
        df_fila = df_filtrado[df_filtrado['Fila Principal'] == fila]
        with cols[i % 4]:
            desenhar_card(fila, df_fila)
            st.markdown("<br>", unsafe_allow_html=True)

else:
    # VISÃO 2: TELA DE DETALHES
    fila_atual = st.session_state.fila_selecionada
    
    st.markdown("""
        <style>
        .btn-voltar-container .stButton>button { border-radius: 6px !important; background-color: #0056b3 !important; color: white !important; margin-top: 0px !important; width: 280px !important; border: none !important; }
        .btn-voltar-container .stButton>button:hover { background-color: #004494 !important; }
        </style>
        <div class="btn-voltar-container">
    """, unsafe_allow_html=True)
    
    if st.button("⬅️ Voltar para a Grade Principal"):
        st.session_state.fila_selecionada = None
        st.rerun()
        
    st.markdown('</div>', unsafe_allow_html=True)
    st.markdown(f"<h2 style='color: #0c1c2b; margin-top: 15px; margin-bottom: 20px;'>Fila: {fila_atual}</h2>", unsafe_allow_html=True)
    
    df_view = df_filtrado[df_filtrado['Fila Principal'] == fila_atual].copy()
    
    # --- SISTEMA DE FILTROS DINÂMICOS NA VISÃO DE DETALHE ---
    col_f1, col_f2, col_f3 = st.columns([2, 2, 4])
    
    with col_f1:
        if fila_atual in ["CORPORATIVO", "ATRIBUÍDO AO USUÁRIO", "SAFETY", "CASOS SEM FILA - GENÉRICO"]:
            label_filtro = "📌 Filtrar Carteira:" if fila_atual == "CORPORATIVO" else "👤 Filtrar Subfila/Usuário:"
            subfilas_disp = sorted(df_view['Subfila'].dropna().unique().tolist())
            subfila_sel = st.selectbox(label_filtro, ["Todos"] + subfilas_disp)
            
            if subfila_sel != "Todos":
                df_view = df_view[df_view['Subfila'] == subfila_sel]
        else:
            st.empty() 

    with col_f2:
        status_disp = sorted(df_view['Status'].dropna().unique().tolist())
        status_sel = st.selectbox("🚥 Filtrar Status:", ["Todos"] + status_disp)
        
        if status_sel != "Todos":
            df_view = df_view[df_view['Status'] == status_sel]
            
    st.markdown("<br>", unsafe_allow_html=True)
            
    # --- SISTEMA DE ABAS ---
    tab1, tab2 = st.tabs(["📊 Indicadores Operacionais", "🛠️ Ações em Massa & Extrato"])

    # === ABA 1: INDICADORES ===
    with tab1:
        vol = len(df_view)
        trat = len(df_view[df_view['Macro Status'] == '🟡 Em Tratativa'])
        fech = len(df_view[df_view['Macro Status'] == '🟢 Fechado'])
        df_fechados = df_view[df_view['Macro Status'] == '🟢 Fechado']
        
        if not df_fechados.empty:
            tma_dias = (df_fechados['Fechamento'] - df_fechados['Abertura']).dt.total_seconds().mean() / (24 * 3600)
            tma_str = f"{tma_dias:.1f} dias"
        else:
            tma_str = "N/A"

        st.markdown(f"""
        <div style="display: flex; gap: 15px; margin-bottom: 25px;">
            <div style="flex: 1; background: #f8f9fa; padding: 15px; border-radius: 8px; border-left: 4px solid #0056b3; box-shadow: 0 2px 4px rgba(0,0,0,0.05);">
                <div style="font-size: 11px; color: #6c757d; text-transform: uppercase; font-weight: bold;">Total de Casos</div>
                <div style="font-size: 22px; color: #0c1c2b; font-weight: bold;">{vol}</div>
            </div>
            <div style="flex: 1; background: #f8f9fa; padding: 15px; border-radius: 8px; border-left: 4px solid #f0ad4e; box-shadow: 0 2px 4px rgba(0,0,0,0.05);">
                <div style="font-size: 11px; color: #6c757d; text-transform: uppercase; font-weight: bold;">Em Tratativa</div>
                <div style="font-size: 22px; color: #0c1c2b; font-weight: bold;">{trat}</div>
            </div>
            <div style="flex: 1; background: #f8f9fa; padding: 15px; border-radius: 8px; border-left: 4px solid #00CC96; box-shadow: 0 2px 4px rgba(0,0,0,0.05);">
                <div style="font-size: 11px; color: #6c757d; text-transform: uppercase; font-weight: bold;">Fechados</div>
                <div style="font-size: 22px; color: #0c1c2b; font-weight: bold;">{fech}</div>
            </div>
            <div style="flex: 1; background: #f8f9fa; padding: 15px; border-radius: 8px; border-left: 4px solid #6f42c1; box-shadow: 0 2px 4px rgba(0,0,0,0.05);">
                <div style="font-size: 11px; color: #6c757d; text-transform: uppercase; font-weight: bold;">TMA (Casos Fechados)</div>
                <div style="font-size: 22px; color: #0c1c2b; font-weight: bold;">{tma_str}</div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        c1, c2 = st.columns(2, gap="large")
        with c1:
            df_abertos = df_view[df_view['Macro Status'] == '🟡 Em Tratativa'].copy()
            if not df_abertos.empty:
                df_abertos['Idade'] = (datetime.now(timezone.utc).replace(tzinfo=None) - df_abertos['Abertura']).dt.days
                bins = [-1, 3, 7, 10000]
                labels = ['0 a 3 Dias', '4 a 7 Dias', '+7 Dias']
                df_abertos['Faixa'] = pd.cut(df_abertos['Idade'], bins=bins, labels=labels)
                df_age = df_abertos['Faixa'].value_counts().reindex(labels).reset_index()
                
                fig_age = px.bar(df_age, x='Faixa', y='count', title='Idade dos Casos Abertos', text='count', color='Faixa',
                                 color_discrete_map={'0 a 3 Dias':'#0056b3', '4 a 7 Dias':'#f0ad4e', '+7 Dias':'#d9534f'})
                fig_age.update_traces(textposition='outside', showlegend=False)
                fig_age.update_layout(height=300, margin=dict(l=0, r=0, t=40, b=0), plot_bgcolor='rgba(0,0,0,0)')
                st.plotly_chart(fig_age, use_container_width=True)
            else:
                st.info("Nenhum caso em aberto no momento para o filtro selecionado.")
                
        with c2:
            df_sla = df_view['SLA Atrasado'].value_counts().reset_index()
            if not df_sla.empty:
                fig_sla = px.pie(df_sla, names='SLA Atrasado', values='count', hole=0.5, title='Saúde do SLA (Total)', 
                                 color='SLA Atrasado', color_discrete_map={'✅ No Prazo':'#00CC96', '🔴 Atrasado':'#EF553B'})
                fig_sla.update_layout(height=300, margin=dict(l=0, r=0, t=40, b=0), plot_bgcolor='rgba(0,0,0,0)')
                st.plotly_chart(fig_sla, use_container_width=True)
            else:
                st.info("Nenhum dado de SLA para exibir neste filtro.")

    # === ABA 2: EXTRATO E AÇÕES EM MASSA ===
    with tab2:
        colunas_ordem_ideal = [
            'Número', 'Status', 'Motivo', 'Link Salesforce', 'Descrição', 
            'Fila Principal', 'Subfila', 'Macro Status', 'Idade (Dias)', 'SLA Atrasado',
            'Origem', 'Tipo Solicitação', 'Conta', 'Abertura', 'Fechamento', 'ID do Caso', 'ID do Proprietário'
        ]
        df_view = df_view[colunas_ordem_ideal]
        df_view.insert(0, 'Selecionar', False)
        
        col_dl, col_aviso = st.columns([1, 3])
        with col_dl:
            def to_excel(df_export):
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_temp = df_export.drop(columns=['Selecionar', 'ID do Proprietário', 'ID do Caso']).copy() 
                    df_temp['Abertura'] = df_temp['Abertura'].dt.tz_localize(None)
                    if df_temp['Fechamento'].notna().any():
                        df_temp['Fechamento'] = df_temp['Fechamento'].dt.tz_localize(None)
                    df_temp.to_excel(writer, index=False, sheet_name='Extrato')
                return output.getvalue()

            st.download_button(
                label=f"📥 Baixar Extrato ({len(df_view)} registros)",
                data=to_excel(df_view),
                file_name=f'extrato_{fila_atual.replace(" ", "_").lower()}.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
        with col_aviso:
            st.caption("💡 **Dica:** Marque a caixa para realizar Ações em Massa abaixo. Dê dois cliques no **Status/Motivo** para editar a tabela e salvar.")

        def colorir_linha(row):
            return ['background-color: #ffebee' if row['SLA Atrasado'] == '🔴 Atrasado' else 'background-color: #ffffff' for _ in row]

        colunas_bloqueadas = df_view.columns.drop(['Selecionar', 'Status', 'Motivo']).tolist()

        edited_df = st.data_editor(
            df_view.style.apply(colorir_linha, axis=1),
            column_config={
                "Selecionar": st.column_config.CheckboxColumn("Selecionar", default=False),
                "ID do Caso": None, 
                "ID do Proprietário": None, 
                "Link Salesforce": st.column_config.LinkColumn("Acessar", display_text="Abrir"),
                "Idade (Dias)": st.column_config.NumberColumn("Idade (Dias)", format="%d"),
                "Abertura": st.column_config.DatetimeColumn("Abertura", format="DD/MM/YYYY HH:mm"),
                "Fechamento": st.column_config.DatetimeColumn("Fechamento", format="DD/MM/YYYY HH:mm"),
                "Descrição": st.column_config.TextColumn("Descrição", width="large")
            },
            disabled=colunas_bloqueadas,
            use_container_width=True,
            hide_index=True,
            key=f"editor_{fila_atual}"
        )
        
        # --- SALVAR ALTERAÇÕES DA TABELA ---
        df_alteracoes = edited_df[(edited_df['Status'] != df_view['Status']) | (edited_df['Motivo'] != df_view['Motivo'])]
        if not df_alteracoes.empty:
            st.warning(f"⚠️ Você alterou {len(df_alteracoes)} linha(s) na tabela.")
            if st.button("💾 Salvar Alterações no Salesforce", type="primary"):
                if api_user_id is None:
                    st.error("Erro: Não foi possível identificar seu usuário para realizar a alteração.")
                else:
                    with st.spinner("Atualizando registros..."):
                        try:
                            for _, row in df_alteracoes.iterrows():
                                id_caso = row['ID do Caso']
                                dono_original = row['ID do Proprietário']
                                
                                if dono_original != api_user_id:
                                    sf.Case.update(id_caso, {'OwnerId': api_user_id}, headers={'Sforce-Auto-Assign': 'FALSE'})
                                    
                                sf.Case.update(id_caso, {'Status': row['Status'], 'FOZ_Motivo__c': row['Motivo']}, headers={'Sforce-Auto-Assign': 'FALSE'})
                                
                                if dono_original != api_user_id:
                                    sf.Case.update(id_caso, {'OwnerId': dono_original}, headers={'Sforce-Auto-Assign': 'FALSE'})
                                
                            st.toast("✅ Alterações salvas com sucesso no Salesforce!")
                            time.sleep(1.5)
                            st.cache_data.clear()
                            st.rerun()
                        except Exception as e:
                            st.error(f"Erro ao salvar alterações: {e}")

        st.markdown("---")

        # --- AÇÕES EM MASSA (TRANSFERIR, COMENTAR, FOLLOW-UP) ---
        casos_selecionados = edited_df[edited_df['Selecionar'] == True]
        if not casos_selecionados.empty:
            st.markdown(f"**{len(casos_selecionados)} caso(s) selecionado(s) para ações em massa:**")
            
            c_transf, c_coment, c_followup = st.columns(3, gap="large")
            
            # --- 1. TRANSFERIR ---
            with c_transf:
                st.markdown("##### 🔄 Transferir Casos")
                dono_selecionado = st.selectbox("Selecione o Novo Proprietário:", [""] + list(lista_proprietarios.keys()))
                
                if st.button("Confirmar Transferência", use_container_width=True):
                    if not dono_selecionado:
                        st.warning("Por favor, selecione um proprietário.")
                    elif api_user_id is None:
                        st.error("Erro: Não foi possível identificar seu usuário.")
                    else:
                        with st.spinner("Transferindo casos no Salesforce..."):
                            novo_id = lista_proprietarios[dono_selecionado]
                            sucessos = 0
                            erros = []
                            
                            for _, row in casos_selecionados.iterrows():
                                id_caso = row['ID do Caso']
                                num_caso = row['Número']
                                dono_original = row['ID do Proprietário']
                                
                                try:
                                    if dono_original != api_user_id:
                                        sf.Case.update(id_caso, {'OwnerId': api_user_id}, headers={'Sforce-Auto-Assign': 'FALSE'})
                                        
                                    if novo_id != api_user_id:
                                        sf.Case.update(id_caso, {'OwnerId': novo_id}, headers={'Sforce-Auto-Assign': 'FALSE'})
                                        
                                    sucessos += 1
                                except Exception as e:
                                    msg_erro = str(e)
                                    try:
                                        import json
                                        msg_erro = json.loads(e.content)[0].get('message', msg_erro)
                                    except:
                                        pass
                                    erros.append(f"Caso {num_caso}: {msg_erro}")
                            
                            if erros:
                                st.error(f"⚠️ O Salesforce bloqueou a transferência de {len(erros)} caso(s):")
                                for err in erros: st.warning(err)
                            if sucessos > 0:
                                st.toast(f"✅ {sucessos} caso(s) transferido(s) com sucesso!")
                                time.sleep(1.5)
                                st.cache_data.clear() 
                                st.rerun() 

            # --- 2. COMENTAR ---
            with c_coment:
                st.markdown("##### 💬 Comentar em Lote")
                novo_comentario = st.text_area("Digite o comentário a ser inserido:", height=68)
                
                if st.button("Inserir Comentário", use_container_width=True):
                    if novo_comentario.strip():
                        with st.spinner("Enviando comentários..."):
                            try:
                                payload = [{'ParentId': row['ID do Caso'], 'CommentBody': novo_comentario} for _, row in casos_selecionados.iterrows()]
                                sf.bulk.CaseComment.insert(payload)
                                st.toast("✅ Comentários inseridos com sucesso!")
                                time.sleep(1.5)
                                st.cache_data.clear()
                                st.rerun()
                            except Exception as e:
                                st.error(f"Erro ao inserir comentário: {e}")
                    else:
                        st.warning("O comentário não pode estar vazio.")

            # --- 3. CRIAR FOLLOW-UP ---
            with c_followup:
                st.markdown("##### 🔔 Criar Follow-up")
                user_followup = st.selectbox("Notificar / Atribuir para:", [""] + list(lista_proprietarios.keys()), key="fup_user")
                descricao_followup = st.text_area("Comentário / Descrição da Tarefa:", height=68)
                
                if st.button("Criar Tarefa", use_container_width=True):
                    if not user_followup or not descricao_followup.strip():
                        st.warning("Selecione um usuário e digite a descrição da tarefa.")
                    else:
                        with st.spinner("Criando tarefas de follow-up..."):
                            try:
                                dono_id_tarefa = lista_proprietarios[user_followup]
                                payload = []
                                for _, row in casos_selecionados.iterrows():
                                    payload.append({
                                        'WhatId': row['ID do Caso'],
                                        'OwnerId': dono_id_tarefa,
                                        'Subject': 'Ação Requerida',
                                        'Description': descricao_followup,
                                        'Status': 'Open',
                                        'Priority': 'Normal'
                                    })
                                sf.bulk.Task.insert(payload)
                                st.toast("✅ Tarefas de Follow-up criadas com sucesso!")
                                time.sleep(1.5)
                                st.cache_data.clear()
                                st.rerun()
                            except Exception as e:
                                st.error(f"Erro ao criar Follow-up: {e}")
