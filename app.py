import streamlit as st
import pandas as pd
from simple_salesforce import Salesforce
from io import BytesIO
from datetime import datetime, timedelta, timezone
import plotly.express as px
import time
import os

# Tenta importar o autorefresh para o Modo TV
try:
    from streamlit_autorefresh import st_autorefresh
    HAS_AUTOREFRESH = True
except ImportError:
    HAS_AUTOREFRESH = False

# --- CONFIGURAÇÃO DE CAMPOS DO SALESFORCE ---
# Alterado para buscar o código do item diretamente no objeto
CAMPO_ITEM_CONTRATO = 'FOZ_CodigoItem__c'

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Gestão de Casos", layout="wide", initial_sidebar_state="expanded")

# --- CSS CUSTOMIZADO ---
st.markdown("""
    <style>
    .stApp { background-color: #ffffff !important; }
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
    .stButton>button:hover { background-color: #0056b3 !important; color: white !important; }
    .btn-login>button {
        border-radius: 8px !important; margin-top: 10px !important;
        background-color: #0056b3 !important; color: white !important; border: none !important;
    }
    .btn-login>button:hover { background-color: #004494 !important; }
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

# --- LEITURA DA BASECORP (EXCEL) ---
@st.cache_data(ttl=3600)
def load_basecorp():
    try:
        # Lê o arquivo Excel local do repositório
        df = pd.read_excel('basecorp.xlsx')
        df.columns = df.columns.str.lower().str.strip()
        
        # Garante o formato texto, tira zeros invisíveis (ex: .0 do Excel) e limpa zeros à esquerda
        df['itemcontrato'] = df['itemcontrato'].astype(str).str.replace('\.0$', '', regex=True).str.strip().str.lstrip('0')
        df['carteira'] = df['carteira'].astype(str).str.strip()
        
        # Transforma em dicionário { '24560': 'Carteira A', '12345': 'Carteira B' }
        return dict(zip(df['itemcontrato'], df['carteira']))
    except Exception as e:
        return {}

# Função auxiliar para extrair campo dinâmico
def extract_field(record, field_path):
    parts = field_path.split('.')
    val = record
    for part in parts:
        if val and isinstance(val, dict):
            val = val.get(part)
        else:
            return None
    return str(val) if val else ""

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
    basecorp_dict = load_basecorp()
    
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
        {CAMPO_ITEM_CONTRATO},
        (SELECT IsViolated, TargetDate FROM CaseMilestones ORDER BY TargetDate ASC),
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
    
    sf_base_url = "https://ibbl.lightning.force.com/lightning/r/Case/"
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
            
            sla_macro = "✅ No Prazo"
            sla_visual = "⚪ Sem SLA"
            sla_atrasado_bool = False
            
            if record['CaseMilestones'] and record['CaseMilestones'].get('records'):
                milestones = record['CaseMilestones']['records']
                sla_atrasado_bool = any(m.get('IsViolated') for m in milestones)
                sla_macro = "🔴 Atrasado" if sla_atrasado_bool else "✅ No Prazo"
                
                target_date_str = milestones[0].get('TargetDate')
                if target_date_str and macro_status != "🟢 Fechado":
                    target_dt = pd.to_datetime(target_date_str).replace(tzinfo=None)
                    diferenca = (target_dt - hoje_utc).days
                    
                    if sla_atrasado_bool or diferenca < 0:
                        sla_visual = f"🔴 Atrasado ({abs(diferenca)} d)"
                    elif diferenca == 0:
                        sla_visual = "🟡 Vence Hoje"
                    else:
                        sla_visual = f"🟢 No Prazo ({diferenca} d)"
                elif macro_status == "🟢 Fechado":
                    sla_visual = "✅ Fechado"
            
            data_abertura = pd.to_datetime(record['CreatedDate']).tz_localize(None)
            data_fechamento = pd.to_datetime(record['ClosedDate']).tz_localize(None) if record.get('ClosedDate') else None
            
            fim_calc = data_fechamento if data_fechamento else hoje_utc
            idade_dias = (fim_calc - data_abertura).days
            
            classificacao = '-'
            if record['Account'] and record['Account'].get('FOZ_Classificacao__c'):
                classificacao = record['Account']['FOZ_Classificacao__c']
                
            # --- INTEGRAÇÃO COM BASECORP ---
            raw_item_contrato = extract_field(record, CAMPO_ITEM_CONTRATO).strip()
            
            # Remove os zeros à esquerda APENAS para buscar no dicionário do Excel
            item_contrato_limpo = raw_item_contrato.lstrip('0') if raw_item_contrato else ""
            if raw_item_contrato and not item_contrato_limpo: # Caso fosse '000000'
                item_contrato_limpo = "0"
                
            carteira_basecorp = basecorp_dict.get(item_contrato_limpo, "-") if item_contrato_limpo else "-"
                
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
                'Abertura': data_abertura,
                'Fechamento': data_fechamento,
                'Origem': record['Origin'],
                'Tipo Solicitação': record['FOZ_TipoSolicitacao__c'],
                'Motivo': record['FOZ_Motivo__c'],
                'SLA (Prazo)': sla_visual,
                'Status': record['Status'],
                'BaseCorp Carteira': carteira_basecorp,
                'Item de Contrato': raw_item_contrato, # Mantém o original com zeros na tabela visual
                'Descrição': descricao_final,
                'Fila Principal': fila_principal,
                'Subfila': subfila,
                'Macro Status': macro_status,
                'SLA Macro': sla_macro,
                'Idade (Dias)': idade_dias,
                'Conta': record['Account']['Name'] if record['Account'] else '-'
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

# --- FUNÇÕES DE MODAIS (POP-UPS) ---
@st.dialog("🔄 Transferir e Comentar")
def modal_transferir_comentar(casos_selecionados_df, lista_prop, api_usr_id):
    st.markdown(f"Você está executando ação em **{len(casos_selecionados_df)} caso(s)**.")
    
    dono_selecionado = st.selectbox("Selecione o Novo Proprietário (*Obrigatório*):", [""] + list(lista_prop.keys()))
    novo_comentario = st.text_area("Adicionar Comentário:", placeholder="(Opcional) Deixe em branco se quiser apenas transferir...", height=100)
    
    if st.button("Confirmar Ação", type="primary", use_container_width=True):
        if not dono_selecionado:
            st.warning("⚠️ Por favor, selecione um proprietário para transferir.")
            return
        if not api_usr_id:
            st.error("⚠️ Erro: Não foi possível identificar seu usuário da API.")
            return
            
        with st.spinner("Processando atualizações no Salesforce..."):
            novo_id = lista_prop[dono_selecionado]
            sucessos = 0
            erros = []
            
            for _, row in casos_selecionados_df.iterrows():
                id_caso = row['ID do Caso']
                num_caso = row['Número']
                dono_original = row['ID do Proprietário']
                
                try:
                    if dono_original != api_usr_id:
                        sf.Case.update(id_caso, {'OwnerId': api_usr_id}, headers={'Sforce-Auto-Assign': 'FALSE'})
                        
                    if novo_id != api_usr_id:
                        sf.Case.update(id_caso, {'OwnerId': novo_id}, headers={'Sforce-Auto-Assign': 'FALSE'})
                        
                    sucessos += 1
                except Exception as e:
                    erros.append(f"Transferência bloqueada no caso {num_caso}: {str(e)}")
            
            if novo_comentario.strip() and sucessos > 0:
                try:
                    payload = [{'ParentId': row['ID do Caso'], 'CommentBody': novo_comentario} for _, row in casos_selecionados_df.iterrows()]
                    sf.bulk.CaseComment.insert(payload)
                except Exception as e:
                    erros.append(f"Erro ao inserir comentários: {str(e)}")
            
            if erros:
                for err in erros: st.error(err)
            if sucessos > 0:
                st.toast(f"✅ {sucessos} ação(ões) concluída(s) com sucesso!")
                time.sleep(1.5)
                st.cache_data.clear()
                st.rerun()

@st.dialog("🔔 Criar Tarefa de Follow-up")
def modal_followup(casos_selecionados_df, lista_prop):
    st.markdown(f"Criando follow-up para **{len(casos_selecionados_df)} caso(s)**.")
    
    user_followup = st.selectbox("Notificar / Atribuir Tarefa para:", [""] + list(lista_prop.keys()))
    descricao_followup = st.text_area("Comentário / Descrição da Tarefa:", height=100)
    
    if st.button("Confirmar Follow-up", type="primary", use_container_width=True):
        if not user_followup or not descricao_followup.strip():
            st.warning("⚠️ Selecione um usuário e digite a descrição da tarefa.")
            return
            
        with st.spinner("Registrando tarefas no Salesforce..."):
            try:
                dono_id_tarefa = lista_prop[user_followup]
                payload = []
                for _, row in casos_selecionados_df.iterrows():
                    payload.append({
                        'WhatId': row['ID do Caso'],
                        'OwnerId': dono_id_tarefa,
                        'Subject': 'Ação Requerida',
                        'Description': descricao_followup,
                        'Status': 'Open',
                        'Priority': 'Normal'
                    })
                sf.bulk.Task.insert(payload)
                st.toast("✅ Tarefas de Follow-up criadas e notificadas com sucesso!")
                time.sleep(1.5)
                st.cache_data.clear()
                st.rerun()
            except Exception as e:
                st.error(f"Erro ao criar Follow-up: {e}")

@st.dialog("🎯 Transferência Inteligente (BaseCorp)")
def modal_basecorp(casos_selecionados_df, lista_prop, api_usr_id):
    st.markdown("O sistema vai ler a planilha **basecorp.xlsx** e transferir automaticamente os casos que possuírem mapeamento.")
    
    acoes_planejadas = []
    sem_mapeamento = []
    
    for _, row in casos_selecionados_df.iterrows():
        carteira_basecorp = row['BaseCorp Carteira']
        
        if carteira_basecorp == "-":
            sem_mapeamento.append(row['Número'])
            continue
            
        novo_id = None
        for key, val in lista_prop.items():
            key_clean = key.replace('📁', '').replace('👤', '').strip().upper()
            if carteira_basecorp.strip().upper() in key_clean:
                novo_id = val
                break
                
        if novo_id:
            acoes_planejadas.append((row, novo_id, carteira_basecorp))
        else:
            sem_mapeamento.append(f"{row['Número']} (Fila '{carteira_basecorp}' não achada no SF)")
            
    if acoes_planejadas:
        st.success(f"**{len(acoes_planejadas)} caso(s)** prontos para roteamento automático.")
    if sem_mapeamento:
        st.warning(f"**{len(sem_mapeamento)} caso(s)** ignorados por falta de mapeamento ou fila inexistente.")
        
    if st.button("Executar Roteamento Automático", type="primary", use_container_width=True):
        if not acoes_planejadas:
            st.error("Nenhum caso apto para transferência automática.")
            return
            
        with st.spinner("Processando atualizações no Salesforce..."):
            sucessos = 0
            erros = []
            
            for row, novo_id, _ in acoes_planejadas:
                id_caso = row['ID do Caso']
                dono_original = row['ID do Proprietário']
                try:
                    if dono_original != api_usr_id:
                        sf.Case.update(id_caso, {'OwnerId': api_usr_id}, headers={'Sforce-Auto-Assign': 'FALSE'})
                    if novo_id != api_usr_id:
                        sf.Case.update(id_caso, {'OwnerId': novo_id}, headers={'Sforce-Auto-Assign': 'FALSE'})
                    sucessos += 1
                except Exception as e:
                    erros.append(str(e))
                    
            if erros:
                st.error("Ocorreram erros em algumas transferências.")
            if sucessos > 0:
                st.toast(f"✅ {sucessos} caso(s) roteado(s) automaticamente com sucesso!")
                time.sleep(1.5)
                st.cache_data.clear()
                st.rerun()

# --- FUNÇÃO PARA DESENHAR O CARD QUADRADO ---
def desenhar_card(fila_nome, df_fila):
    vol = len(df_fila)
    trat = len(df_fila[df_fila['Macro Status'] == '🟡 Em Tratativa'])
    fech = len(df_fila[df_fila['Macro Status'] == '🟢 Fechado'])
    atr = len(df_fila[(df_fila['SLA Macro'] == '🔴 Atrasado') & (df_fila['Macro Status'] == '🟡 Em Tratativa')])
    
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

if HAS_AUTOREFRESH:
    modo_tv = st.sidebar.toggle("⏱️ Modo TV (Atualiza a cada 5 min)")
    if modo_tv:
        st_autorefresh(interval=5 * 60 * 1000, key="data_refresh")
else:
    st.sidebar.caption("💡 Para habilitar o Modo TV (Auto-Refresh), instale o pacote via terminal: `pip install streamlit-autorefresh`")

st.sidebar.markdown("---")
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

filtro_meus_casos = st.sidebar.toggle("🙋‍♂️ Ver apenas Meus Casos", value=False)
incluir_fechados = st.sidebar.checkbox("Mostrar Casos Fechados", value=False)

user, pwd, token = st.session_state.sf_username, st.session_state.sf_password, st.session_state.sf_token
df_filtrado = get_data(periodo_selecionado, dt_inicio, dt_fim, incluir_fechados, user, pwd, token)
lista_proprietarios = get_owner_options(user, pwd, token)
api_user_id = get_api_user_id(user, pwd, token)

if filtro_meus_casos and not df_filtrado.empty and api_user_id:
    df_filtrado = df_filtrado[df_filtrado['ID do Proprietário'] == api_user_id]

if busca_global and not df_filtrado.empty:
    mask = (
        df_filtrado['Número'].astype(str).str.contains(busca_global, case=False, na=False) |
        df_filtrado['Conta'].astype(str).str.contains(busca_global, case=False, na=False)
    )
    df_filtrado = df_filtrado[mask]

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
            
    tab1, tab2 = st.tabs(["📊 Indicadores Operacionais", "🛠️ Ações em Massa & Extrato"])

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
            df_sla = df_view['SLA Macro'].value_counts().reset_index()
            if not df_sla.empty:
                fig_sla = px.pie(df_sla, names='SLA Macro', values='count', hole=0.5, title='Saúde do SLA (Total)', 
                                 color='SLA Macro', color_discrete_map={'✅ No Prazo':'#00CC96', '🔴 Atrasado':'#EF553B'})
                fig_sla.update_layout(height=300, margin=dict(l=0, r=0, t=40, b=0), plot_bgcolor='rgba(0,0,0,0)')
                st.plotly_chart(fig_sla, use_container_width=True)
            else:
                st.info("Nenhum dado de SLA para exibir neste filtro.")

    with tab2:
        colunas_ordem_ideal = [
            'Número', 'Abertura', 'Fechamento', 'Origem', 'Tipo Solicitação', 'Motivo',
            'SLA (Prazo)', 'Status', 'BaseCorp Carteira', 'Item de Contrato', 'Link Salesforce', 'Descrição', 
            'Fila Principal', 'Subfila', 'Macro Status', 'Idade (Dias)', 'Conta', 'ID do Caso', 'ID do Proprietário'
        ]
        df_view = df_view[colunas_ordem_ideal]
        df_view.insert(0, 'Selecionar', False)
        
        container_acoes = st.container()
        st.markdown("<br>", unsafe_allow_html=True)
        container_tabela = st.container()
        container_rodape = st.container()
        
        with container_tabela:
            st.caption("💡 **Dica:** Marque a caixa 'Selecionar' na tabela para exibir as **Ações em Massa** no topo. Dê dois cliques no **Status/Motivo** para editar e salvar.")

            def colorir_linha(row):
                return ['background-color: #ffebee' if 'Atrasado' in row['SLA (Prazo)'] else 'background-color: #ffffff' for _ in row]

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
        
        casos_selecionados = edited_df[edited_df['Selecionar'] == True]
        
        with container_acoes:
            if not casos_selecionados.empty:
                st.markdown(f"**⚡ Ações Disponíveis para {len(casos_selecionados)} caso(s) selecionado(s):**")
                c_btn1, c_btn2, c_btn3, c_btn4 = st.columns([1.2, 1, 1, 2])
                
                with c_btn1:
                    if st.button("🔄 Transferir e Comentar", use_container_width=True):
                        modal_transferir_comentar(casos_selecionados, lista_proprietarios, api_user_id)
                with c_btn2:
                    if st.button("🎯 Roteamento BaseCorp", use_container_width=True):
                        modal_basecorp(casos_selecionados, lista_proprietarios, api_user_id)
                with c_btn3:
                    if st.button("🔔 Criar Follow-up", use_container_width=True):
                        modal_followup(casos_selecionados, lista_proprietarios)

        with container_rodape:
            st.markdown("---")
            
            df_alteracoes = edited_df[(edited_df['Status'] != df_view['Status']) | (edited_df['Motivo'] != df_view['Motivo'])]
            if not df_alteracoes.empty:
                st.warning(f"⚠️ Você alterou {len(df_alteracoes)} linha(s) na tabela.")
                if st.button("💾 Salvar Alterações no Salesforce", type="primary"):
                    if api_user_id is None:
                        st.error("Erro: Não foi possível identificar seu usuário para realizar a alteração.")
                    else:
                        with st.spinner("Processando atualizações no Salesforce..."):
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

            st.markdown("<br>", unsafe_allow_html=True)
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
