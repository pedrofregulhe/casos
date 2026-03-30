import streamlit as st
import pandas as pd
from simple_salesforce import Salesforce
from io import BytesIO
from datetime import datetime, timedelta, timezone
import plotly.express as px
import time

# Tenta importar o autorefresh para a Atualização Automática
try:
    from streamlit_autorefresh import st_autorefresh
    HAS_AUTOREFRESH = True
except ImportError:
    HAS_AUTOREFRESH = False

# --- CONFIGURAÇÃO DE CAMPOS DO SALESFORCE ---
CAMPO_ITEM_CONTRATO = 'FOZ_Asset__r.FOZ_CodigoItem__c'

# --- CREDENCIAIS SEGURAS (SECRETS) ---
try:
    SF_USER = st.secrets["SF_USERNAME"]
    SF_PWD = st.secrets["SF_PASSWORD"]
    SF_TOKEN = st.secrets["SF_TOKEN"]
    
    APP_USER = st.secrets["APP_USER"]
    APP_PWD = st.secrets["APP_PWD"]
except Exception:
    st.error("⚠️ Configuração de Secrets ausente. Configure SF_USERNAME, SF_PASSWORD, SF_TOKEN, APP_USER e APP_PWD no Streamlit Cloud.")
    st.stop()

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Gestão de Casos (OA & OS)", layout="wide", initial_sidebar_state="expanded")

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
if 'app_authenticated' not in st.session_state:
    st.session_state.app_authenticated = False
if 'fila_selecionada' not in st.session_state:
    st.session_state.fila_selecionada = None
if 'last_update' not in st.session_state:
    st.session_state.last_update = datetime.now(fuso_br).strftime("%d/%m/%Y %H:%M")

# --- TELA DE LOGIN DO PAINEL ---
if not st.session_state.app_authenticated:
    col_vazia1, col_login, col_vazia2 = st.columns([1, 2, 1])
    with col_login:
        st.markdown("<h2 style='text-align: center; color: #0c1c2b; margin-top: 50px;'>🔐 Acesso ao Painel de Casos</h2>", unsafe_allow_html=True)
        st.markdown("<p style='text-align: center; color: #6c757d; margin-bottom: 30px;'>Insira suas credenciais de acesso ao sistema.</p>", unsafe_allow_html=True)
        
        with st.form("login_form"):
            user_input = st.text_input("👤 Usuário")
            pwd_input = st.text_input("🔑 Senha", type="password")
            
            st.markdown("<div class='btn-login'>", unsafe_allow_html=True)
            submitted = st.form_submit_button("Entrar no Painel", use_container_width=True)
            st.markdown("</div>", unsafe_allow_html=True)
            
            if submitted:
                if user_input == APP_USER and pwd_input == APP_PWD:
                    st.session_state.app_authenticated = True
                    st.rerun()
                else:
                    st.error("❌ Usuário ou senha incorretos.")
    st.stop()

# --- CONEXÃO DINÂMICA COM SALESFORCE ---
@st.cache_resource
def init_connection(user, pwd, token):
    return Salesforce(username=user, password=pwd, security_token=token, domain='login')

try:
    sf = init_connection(SF_USER, SF_PWD, SF_TOKEN)
except Exception as e:
    st.error(f"❌ Falha ao conectar com o Salesforce: {e}")
    st.stop()

# --- LEITURA DA BASECORP (EXCEL) ---
@st.cache_data(ttl=3600)
def load_basecorp():
    try:
        df = pd.read_excel('basecorp.xlsx')
        df.columns = df.columns.str.lower().str.strip()
        df['itemcontrato'] = df['itemcontrato'].astype(str).str.replace('\.0$', '', regex=True).str.strip().str.lstrip('0')
        df['carteira'] = df['carteira'].astype(str).str.strip()
        return dict(zip(df['itemcontrato'], df['carteira']))
    except Exception as e:
        return {}

def extract_field(record, field_path):
    parts = field_path.split('.')
    val = record
    for part in parts:
        if val and isinstance(val, dict):
            val = val.get(part)
        else:
            return ""
    return str(val) if val is not None else ""

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

    # Consulta primária (Casos, Contas, Ativos)
    query = f"""
    SELECT 
        Id, CaseNumber, CreatedDate, ClosedDate, Status, Description, Origin, Type, 
        FOZ_TipoSolicitacao__c, FOZ_Motivo__c, FOZ_Detalhe__c, FOZ_Subdetalhe__c, FOZ_SubStatus__c, OwnerId, Owner.Name, 
        Account.Name, Account.FOZ_CPF__c, Account.FOZ_CNPJ__c, Account.FOZ_Classificacao__c, Account.FOZ_StatusPosicaoFinanceira__c,
        FOZ_Asset__r.Status, FOZ_Asset__r.FOZ_EndFranquiaForm__c, FOZ_Asset__r.InstallDate, {CAMPO_ITEM_CONTRATO},
        (SELECT IsViolated, TargetDate FROM CaseMilestones ORDER BY TargetDate ASC),
        (SELECT CommentBody, CreatedBy.Name, CreatedDate FROM CaseComments ORDER BY CreatedDate ASC)
    FROM Case 
    WHERE (Type IN ('OA', 'OS')
           OR (Owner.Name LIKE 'CARTEIRA%') 
           OR Owner.Name LIKE '%GENÉRICO%'
           OR Owner.Name LIKE '%GENERICO%'
           OR Owner.Name LIKE '%Casos sem fila%')
      AND {filtro_data}
      {filtro_status}
    """
    
    sf_base_url = "https://ibbl.lightning.force.com/lightning/r/Case/"
    my_bar = st.progress(0, text="Iniciando sincronização com o Salesforce...")
    
    try:
        # ETAPA 1: Puxa todos os casos
        result = sf_conn.query(query)
        total_records = result.get('totalSize', 0)
        records = result.get('records', [])
        
        if total_records > 0:
            current_len = len(records)
            percent = int((current_len / total_records) * 30)
            my_bar.progress(percent, text=f"Baixando casos... ({current_len} de {total_records})")

            while not result.get('done', True):
                result = sf_conn.query_more(result['nextRecordsUrl'], True)
                records.extend(result.get('records', []))
                current_len = len(records)
                percent = int((current_len / total_records) * 30)
                my_bar.progress(percent, text=f"Baixando casos... ({current_len} de {total_records})")

        my_bar.progress(35, text="Download de Casos concluído! Buscando Itens de Ordem de Serviço...")

        # ETAPA 2: Consulta Secundária em Lote (Buscando o "Neto" diretamente)
        os_dict = {}
        case_ids = [r['Id'] for r in records]
        
        if case_ids:
            chunk_size = 200 # Salesforce aceita listas curtas em comandos IN
            for i in range(0, len(case_ids), chunk_size):
                chunk = case_ids[i:i+chunk_size]
                ids_str = ",".join([f"'{cid}'" for cid in chunk])
                
                query_os = f"""
                SELECT WorkOrder.CaseId, FOZ_Numero_OS__c, FOZ_Nome_Franquia__c, 
                       FOZ_Tipo_de_Servico__c, FOZ_Agendado_para_data_periodo__c, FOZ_Id_Tecnico__c
                FROM WorkOrderLineItem 
                WHERE WorkOrder.CaseId IN ({ids_str})
                ORDER BY CreatedDate DESC
                """
                try:
                    os_result = sf_conn.query_all(query_os)
                    for os_rec in os_result.get('records', []):
                        c_id = os_rec['WorkOrder']['CaseId']
                        # Como ordenamos por DESC, o primeiro que aparecer é o mais recente
                        if c_id not in os_dict:
                            os_dict[c_id] = os_rec
                except Exception as e:
                    print(f"Aviso de Subquery OS: {e}")

        my_bar.progress(50, text="Processamento inteligente de dados em andamento...")

        linhas = []
        hoje_utc = datetime.now(timezone.utc).replace(tzinfo=None)
        total_processar = len(records)
        
        for i, record in enumerate(records):
            if total_processar > 0 and i % 500 == 0:
                progresso_atual = 50 + int((i / total_processar) * 50)
                my_bar.progress(progresso_atual, text=f"Estruturando dados visuais... {progresso_atual}%")

            dono_upper = record['Owner']['Name'].upper() if record['Owner'] else 'SISTEMA/SEM DONO'
            
            filas_conhecidas = ["ERRO SISTÊMICO", "CAPACIDADE", "FRANQUIAS", "AUDITORIA", "HELP TEC", "JURÍDICO", "INFORMAÇÃO", "RAF", "FINANCEIRO", "BACKOFFICE"]
            
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
            data_abertura = pd.to_datetime(record['CreatedDate']).tz_localize(None)
            data_fechamento = pd.to_datetime(record['ClosedDate']).tz_localize(None) if record.get('ClosedDate') else None
            
            # --- LÓGICA INTELIGENTE DE SLA ---
            sla_macro = "✅ No Prazo"
            sla_visual = "⚪ Sem SLA"
            sla_atrasado_bool = False
            
            status_real_sf = record['Status'].strip().lower() if record['Status'] else ""
            
            if fila_principal in ["CASOS SEM FILA - GENÉRICO", "CORPORATIVO"] and status_real_sf in ["aberto", "em aberto"]:
                target_dt_custom = data_abertura + timedelta(hours=24)
                diferenca_horas = (target_dt_custom - hoje_utc).total_seconds() / 3600
                
                sla_atrasado_bool = diferenca_horas < 0
                sla_macro = "🔴 Atrasado" if sla_atrasado_bool else "✅ No Prazo"
                
                if sla_atrasado_bool:
                    horas_atraso = abs(diferenca_horas)
                    if horas_atraso >= 24:
                        sla_visual = f"🔴 Atrasado ({int(horas_atraso/24)} d)"
                    else:
                        sla_visual = f"🔴 Atrasado ({int(horas_atraso)} h)"
                else:
                    if diferenca_horas <= 4:
                        sla_visual = f"🟡 Vence Hoje ({int(diferenca_horas)} h)"
                    elif diferenca_horas >= 24:
                        sla_visual = f"🟢 No Prazo ({int(diferenca_horas/24)} d)"
                    else:
                        sla_visual = f"🟢 No Prazo ({int(diferenca_horas)} h)"
            else:
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
                elif macro_status == "🟢 Fechado":
                    sla_visual = "✅ Fechado"

            fim_calc = data_fechamento if data_fechamento else hoje_utc
            idade_dias = (fim_calc - data_abertura).days
            
            # --- EXTRAÇÃO DE CONTA E ATIVO ---
            acc = record.get('Account') or {}
            acc_name = acc.get('Name', '-')
            acc_cnpj = acc.get('FOZ_CNPJ__c', '-')
            acc_classificacao = acc.get('FOZ_Classificacao__c', '-')
            acc_posfin = acc.get('FOZ_StatusPosicaoFinanceira__c', '-')
            
            asset = record.get('FOZ_Asset__r') or {}
            asset_status = asset.get('Status', '-')
            asset_end = asset.get('FOZ_EndFranquiaForm__c', '-')
            asset_install = asset.get('InstallDate', '-')
            
            raw_item_contrato = extract_field(record, CAMPO_ITEM_CONTRATO).strip()
            item_contrato_limpo = raw_item_contrato.lstrip('0') if raw_item_contrato else ""
            if raw_item_contrato and not item_contrato_limpo: item_contrato_limpo = "0"
            carteira_basecorp = basecorp_dict.get(item_contrato_limpo, "-") if item_contrato_limpo else "-"
                
            # --- MESCLANDO A ORDEM DE SERVIÇO ---
            case_id = record['Id']
            os_info = os_dict.get(case_id, {})
            os_numero = os_info.get('FOZ_Numero_OS__c', '')
            os_franquia = os_info.get('FOZ_Nome_Franquia__c', '')
            os_tipo_servico = os_info.get('FOZ_Tipo_de_Servico__c', '')
            os_agendamento = os_info.get('FOZ_Agendado_para_data_periodo__c', '')
            os_tecnico = os_info.get('FOZ_Id_Tecnico__c', '')
                
            # --- COMENTÁRIOS E DESCRIÇÃO ---
            desc_oficial = record.get('Description')
            historico_comentarios = ""
            if record.get('CaseComments') and record['CaseComments'].get('records'):
                for comment in record['CaseComments']['records']:
                    autor = comment['CreatedBy']['Name'] if comment.get('CreatedBy') else 'Usuário'
                    try:
                        dt_obj = pd.to_datetime(comment['CreatedDate']).tz_convert(fuso_br)
                        data_str = dt_obj.strftime('%d/%m/%Y %H:%M')
                    except:
                        data_str = comment['CreatedDate']
                    historico_comentarios += f"🗣️ {autor} em {data_str}:\n{comment['CommentBody']}\n\n"
            
            if desc_oficial and historico_comentarios:
                descricao_final = f"📝 DESCRIÇÃO ORIGINAL:\n{desc_oficial}\n\n{'-'*40}\n\n💬 HISTÓRICO DE COMENTÁRIOS:\n{historico_comentarios}".strip()
            elif historico_comentarios:
                descricao_final = f"💬 HISTÓRICO DE COMENTÁRIOS:\n{historico_comentarios}".strip()
            else:
                descricao_final = desc_oficial if desc_oficial else "-"
                
            tipo_caso = str(record.get('Type')).upper()
            tipo_aba = 'OS' if 'OS' in tipo_caso else 'OA'
                
            linhas.append({
                'ID do Caso': record['Id'],
                'ID do Proprietário': record['OwnerId'],
                'Link Salesforce': f"{sf_base_url}{record['Id']}/view",
                'Número': record['CaseNumber'],
                'Tipo Aba': tipo_aba,
                'Abertura': data_abertura,
                'Fechamento': data_fechamento,
                'Origem': record['Origin'],
                'Tipo Solicitação': record['FOZ_TipoSolicitacao__c'],
                'Motivo': record['FOZ_Motivo__c'],
                'Detalhe': record['FOZ_Detalhe__c'],
                'Substatus': record['FOZ_SubStatus__c'] if record['FOZ_SubStatus__c'] else "",
                'SLA (Prazo)': sla_visual,
                'Status': record['Status'],
                'BaseCorp Carteira': carteira_basecorp,
                'Conta': acc_name,
                'Conta - CNPJ': acc_cnpj,
                'Conta - Posição Fin.': acc_posfin,
                'Conta - Classificação': acc_classificacao,
                'Item de Contrato': raw_item_contrato, 
                'Asset - Status': asset_status,
                'Asset - Endereço': asset_end,
                'Asset - Instalação': asset_install,
                'OS - Número': os_numero,
                'OS - Franquia': os_franquia,
                'OS - Tipo Serviço': os_tipo_servico,
                'OS - Agendamento': os_agendamento,
                'OS - Técnico': os_tecnico,
                'Descrição': descricao_final,
                'Fila Principal': fila_principal,
                'Subfila': subfila,
                'Macro Status': macro_status,
                'SLA Macro': sla_macro,
                'Idade (Dias)': idade_dias
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


# --- FUNÇÕES DE MODAIS E AUDITORIA AUTOMÁTICA ---
def criar_comentario_auditoria(num_caso, id_caso, extra_texto=""):
    texto_base = "Caso editado ou movimentado via Automação."
    if extra_texto.strip(): texto_base += f"\n\nObservação: {extra_texto}"
    return {'ParentId': id_caso, 'CommentBody': texto_base}

@st.dialog("📄 Resumo Diário da Operação")
def modal_resumo_diario(df_dados, per_sel, dt_ini, dt_fim):
    if df_dados.empty:
        st.warning("Não há dados para resumir.")
        return
        
    hoje_str = datetime.now(fuso_br).strftime("%d/%m/%Y às %H:%M")
    periodo_txt = f"{dt_ini.strftime('%d/%m/%Y')} a {dt_fim.strftime('%d/%m/%Y')}" if per_sel == "Personalizado" else per_sel
    
    vol_total = len(df_dados)
    trat_total = len(df_dados[df_dados['Macro Status'] == '🟡 Em Tratativa'])
    fech_total = len(df_dados[df_dados['Macro Status'] == '🟢 Fechado'])
    atr_total = len(df_dados[(df_dados['SLA Macro'] == '🔴 Atrasado') & (df_dados['Macro Status'] == '🟡 Em Tratativa')])
    
    resumo = f"📊 *STATUS DA OPERAÇÃO - Atualizado em {hoje_str}*\n"
    resumo += f"📅 *Período Base Consultado:* {periodo_txt}\n\n"
    
    resumo += f"📈 *VISÃO GERAL DA BASE*\n"
    resumo += f"▪️ Total de Casos: {vol_total}\n"
    resumo += f"▪️ Casos Abertos (Em Tratativa): {trat_total}\n"
    resumo += f"▪️ Casos Fechados: {fech_total}\n"
    resumo += f"▪️ SLA Atrasado (Em Aberto): {atr_total}\n\n"
    
    resumo += f"🎯 *DESTAQUES OPERACIONAIS*\n"
    for fila_destaque in ["ATRIBUÍDO AO USUÁRIO", "CORPORATIVO"]:
        df_destaque = df_dados[df_dados['Fila Principal'] == fila_destaque]
        if not df_destaque.empty:
            vol_d = len(df_destaque)
            abertos_d = len(df_destaque[df_destaque['Macro Status'] == '🟡 Em Tratativa'])
            atr_d = len(df_destaque[(df_destaque['SLA Macro'] == '🔴 Atrasado') & (df_destaque['Macro Status'] == '🟡 Em Tratativa')])
            resumo += f"🔸 *{fila_destaque}:* {vol_d} Casos no total | {abertos_d} Abertos | {atr_d} Atrasados\n"
            subfilas = sorted(df_destaque['Subfila'].dropna().unique().tolist())
            for sub in subfilas:
                if sub == "-": continue
                df_sub = df_destaque[df_destaque['Subfila'] == sub]
                resumo += f"   ↳ {sub}: {len(df_sub)} Total | {len(df_sub[df_sub['Macro Status'] == '🟡 Em Tratativa'])} Abertos | {len(df_sub[(df_sub['SLA Macro'] == '🔴 Atrasado') & (df_sub['Macro Status'] == '🟡 Em Tratativa')])} Atrasados\n"
            resumo += "\n"
            
    resumo += f"🏢 *DETALHAMENTO POR FILA GERAL*\n"
    filas = sorted(df_dados['Fila Principal'].dropna().unique().tolist())
    for fila in filas:
        if fila in ["ATRIBUÍDO AO USUÁRIO", "CORPORATIVO"]: continue 
        df_fila = df_dados[df_dados['Fila Principal'] == fila]
        vol = len(df_fila)
        trat = len(df_fila[df_fila['Macro Status'] == '🟡 Em Tratativa'])
        fech = len(df_fila[df_fila['Macro Status'] == '🟢 Fechado'])
        atr = len(df_fila[(df_fila['SLA Macro'] == '🔴 Atrasado') & (df_fila['Macro Status'] == '🟡 Em Tratativa')])
        
        df_fech = df_fila[df_fila['Macro Status'] == '🟢 Fechado']
        tma_str = f"{(df_fech['Fechamento'] - df_fech['Abertura']).dt.total_seconds().mean() / (24 * 3600):.1f} dias" if not df_fech.empty else "N/A"
            
        resumo += f"\n🔹 *{fila}*\n"
        resumo += f"   Total: {vol} | Abertos: {trat} | Fechados: {fech}\n"
        resumo += f"   SLA Atrasado: {atr} | TMA Médio: {tma_str}\n"

    st.markdown("💡 **Clique no ícone de 'Copiar'** no canto superior direito para copiar.")
    st.code(resumo, language="markdown")

@st.dialog("🔄 Transferir e Comentar")
def modal_transferir_comentar(casos_selecionados_df, lista_prop):
    st.markdown(f"Você está transferindo **{len(casos_selecionados_df)} caso(s)**.")
    tem_basecorp = not casos_selecionados_df[casos_selecionados_df['BaseCorp Carteira'] != '-'].empty
    modo_transferencia = "Manual"
    dono_selecionado = None
    
    if tem_basecorp:
        st.info("🎯 Identificamos caso(s) com mapeamento na BaseCorp.")
        modo_transferencia = st.radio("Como deseja realizar a transferência?", ["Manual (Escolher nova fila/usuário)", "Inteligente (Usar roteamento BaseCorp)"])
        
    if modo_transferencia.startswith("Manual"):
        dono_selecionado = st.selectbox("Selecione o Novo Proprietário (*Obrigatório*):", [""] + list(lista_prop.keys()))
        
    novo_comentario = st.text_area("Adicionar Comentário Opcional:", height=68)
    
    st.markdown("---")
    sem_comentario = st.checkbox("Movimentar sem inserir o comentário padrão de automação", value=False)
    senha_input = st.text_input("🔑 Senha do Salesforce (*Obrigatória*)", type="password")
    
    if st.button("Confirmar Transferência", type="primary", use_container_width=True):
        if senha_input != SF_PWD:
            st.error("⚠️ Senha incorreta. Operação cancelada.")
            return
        if modo_transferencia.startswith("Manual") and not dono_selecionado:
            st.warning("⚠️ Selecione um proprietário para transferir.")
            return
            
        with st.spinner("Processando..."):
            sucessos, erros, comentarios_payload = 0, [], []
            for _, row in casos_selecionados_df.iterrows():
                id_caso = row['ID do Caso']
                num_caso = row['Número']
                if row['Status'] in ['Closed', 'Fechado']:
                    erros.append(f"Caso {num_caso} ignorado: Caso fechado.")
                    continue
                
                novo_id = lista_prop.get(dono_selecionado) if modo_transferencia.startswith("Manual") else None
                if not modo_transferencia.startswith("Manual"):
                    carteira_bc = row['BaseCorp Carteira']
                    if carteira_bc != '-':
                        for key, val in lista_prop.items():
                            if carteira_bc.strip().upper() in key.replace('📁', '').replace('👤', '').strip().upper():
                                novo_id = val
                                break
                
                if novo_id:
                    try:
                        sf.Case.update(id_caso, {'OwnerId': novo_id, 'FOZ_Bypass_Flow__c': True}, headers={'Sforce-Auto-Assign': 'FALSE'})
                        sf.Case.update(id_caso, {'FOZ_Bypass_Flow__c': False}, headers={'Sforce-Auto-Assign': 'FALSE'})
                        sucessos += 1
                        if not sem_comentario or novo_comentario.strip():
                            comentarios_payload.append(criar_comentario_auditoria(num_caso, id_caso, novo_comentario) if not sem_comentario else {'ParentId': id_caso, 'CommentBody': novo_comentario})
                    except Exception as e:
                        erros.append(f"Erro no caso {num_caso}: {str(e)}")
            
            if comentarios_payload:
                try: sf.bulk.CaseComment.insert(comentarios_payload)
                except: pass
            
            if erros:
                for err in erros: st.error(err)
            if sucessos > 0:
                st.toast(f"✅ {sucessos} transferido(s)!")
                time.sleep(1.5)
                st.cache_data.clear()
                st.rerun()

@st.dialog("📝 Editar Casos")
def modal_editar_casos(casos_selecionados_df, df_view):
    st.markdown(f"Editando **{len(casos_selecionados_df)} caso(s)**.")
    opcoes_status = sorted(df_view['Status'].dropna().unique().tolist())
    if "Fechado" not in opcoes_status: opcoes_status.append("Fechado")
        
    novo_status = st.selectbox("Novo Status:", opcoes_status, index=None, placeholder="Selecione para alterar...")
    novo_substatus = st.selectbox("Novo Substatus:", ["Sucesso", "Insucesso", "Indevido", "Limpar Campo (Vazio)"], index=None, placeholder="Selecione para alterar...")
    
    st.markdown("---")
    sem_comentario = st.checkbox("Movimentar sem inserir o comentário padrão de automação", value=False)
    senha_input = st.text_input("🔑 Senha do Salesforce (*Obrigatória*)", type="password")
    
    if st.button("💾 Confirmar Edições", type="primary", use_container_width=True):
        if senha_input != SF_PWD:
            st.error("⚠️ Senha incorreta.")
            return
        if not novo_status and not novo_substatus:
            st.warning("Nenhuma alteração selecionada.")
            return
            
        with st.spinner("Editando..."):
            sucessos, erros, comentarios_payload = 0, [], []
            for _, row in casos_selecionados_df.iterrows():
                id_caso, num_caso = row['ID do Caso'], row['Número']
                payload = {'FOZ_Bypass_Flow__c': True}
                if novo_status: payload['Status'] = novo_status
                if novo_substatus: payload['FOZ_SubStatus__c'] = "" if novo_substatus == "Limpar Campo (Vazio)" else novo_substatus
                    
                try:
                    sf.Case.update(id_caso, payload, headers={'Sforce-Auto-Assign': 'FALSE'})
                    sf.Case.update(id_caso, {'FOZ_Bypass_Flow__c': False}, headers={'Sforce-Auto-Assign': 'FALSE'})
                    sucessos += 1
                    if not sem_comentario:
                        comentarios_payload.append(criar_comentario_auditoria(num_caso, id_caso, f"Edição em Lote - Status: {novo_status or 'Mantido'} | Substatus: {novo_substatus or 'Mantido'}"))
                except Exception as e:
                    erros.append(f"Erro no caso {num_caso}: {str(e)}")
            
            if comentarios_payload:
                try: sf.bulk.CaseComment.insert(comentarios_payload)
                except: pass
            
            if erros:
                for err in erros: st.error(err)
            if sucessos > 0:
                st.toast(f"✅ {sucessos} editado(s)!")
                time.sleep(1.5)
                st.cache_data.clear()
                st.rerun()

@st.dialog("🔔 Criar Tarefa de Follow-up")
def modal_followup(casos_selecionados_df, lista_prop):
    st.markdown(f"Criando follow-up para **{len(casos_selecionados_df)} caso(s)**.")
    users_followup = st.multiselect("Notificar / Atribuir Tarefa para (Selecione vários se desejar):", list(lista_prop.keys()))
    descricao_followup = st.text_area("Descrição da Tarefa:", height=100)
    
    st.markdown("---")
    sem_comentario = st.checkbox("Movimentar sem inserir o comentário padrão de automação", value=False)
    senha_input = st.text_input("🔑 Senha do Salesforce (*Obrigatória*)", type="password")
    
    if st.button("Confirmar Follow-up", type="primary", use_container_width=True):
        if senha_input != SF_PWD:
            st.error("⚠️ Senha incorreta.")
            return
        if not users_followup or not descricao_followup.strip():
            st.warning("⚠️ Preencha os campos.")
            return
            
        with st.spinner("Registrando tarefas..."):
            try:
                payload, comentarios_payload = [], []
                for user_nome in users_followup:
                    for _, row in casos_selecionados_df.iterrows():
                        payload.append({'WhatId': row['ID do Caso'], 'OwnerId': lista_prop[user_nome], 'Subject': 'Ação Requerida', 'Description': descricao_followup, 'Status': 'Open', 'Priority': 'Normal'})
                        if not sem_comentario: comentarios_payload.append(criar_comentario_auditoria(row['Número'], row['ID do Caso'], "Nova Tarefa criada."))
                            
                sf.bulk.Task.insert(payload)
                if comentarios_payload:
                    try: sf.bulk.CaseComment.insert(list({c['ParentId']: c for c in comentarios_payload}.values()))
                    except: pass
                st.toast(f"✅ {len(payload)} Tarefas criadas!")
                time.sleep(1.5)
                st.cache_data.clear()
                st.rerun()
            except Exception as e:
                st.error(f"Erro: {e}")

# --- FUNÇÃO GENÉRICA DE RENDERIZAÇÃO DE TELA (ABAS OA/OS) ---
def render_aba_conteudo(df_aba, tipo_aba, fila_atual, lista_proprietarios):
    if df_aba.empty:
        st.info(f"Nenhum caso do tipo **{tipo_aba}** encontrado nesta visão.")
        return

    vol = len(df_aba)
    trat = len(df_aba[df_aba['Macro Status'] == '🟡 Em Tratativa'])
    fech = len(df_aba[df_aba['Macro Status'] == '🟢 Fechado'])
    atr = len(df_aba[(df_aba['SLA Macro'] == '🔴 Atrasado') & (df_aba['Macro Status'] == '🟡 Em Tratativa')])
    tma_str = "N/A"
    df_fechados = df_aba[df_aba['Macro Status'] == '🟢 Fechado']
    if not df_fechados.empty:
        tma_dias = (df_fechados['Fechamento'] - df_fechados['Abertura']).dt.total_seconds().mean() / (24 * 3600)
        tma_str = f"{tma_dias:.1f} dias"

    st.markdown(f"""
    <div style="display: flex; gap: 15px; margin-bottom: 25px;">
        <div style="flex: 1; background: #f8f9fa; padding: 15px; border-radius: 8px; border-left: 4px solid #0056b3; box-shadow: 0 2px 4px rgba(0,0,0,0.05);">
            <div style="font-size: 11px; color: #6c757d; text-transform: uppercase; font-weight: bold;">Total de Casos ({tipo_aba})</div>
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
            <div style="font-size: 11px; color: #6c757d; text-transform: uppercase; font-weight: bold;">TMA Médio</div>
            <div style="font-size: 22px; color: #0c1c2b; font-weight: bold;">{tma_str}</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    c1, c2 = st.columns(2, gap="large")
    with c1:
        df_abertos = df_aba[df_aba['Macro Status'] == '🟡 Em Tratativa'].copy()
        if not df_abertos.empty:
            df_abertos['Idade'] = (datetime.now(timezone.utc).replace(tzinfo=None) - df_abertos['Abertura']).dt.days
            df_abertos['Faixa'] = pd.cut(df_abertos['Idade'], bins=[-1, 3, 7, 10000], labels=['0 a 3 Dias', '4 a 7 Dias', '+7 Dias'])
            df_age = df_abertos['Faixa'].value_counts().reindex(['0 a 3 Dias', '4 a 7 Dias', '+7 Dias']).reset_index()
            fig_age = px.bar(df_age, x='Faixa', y='count', title='Idade dos Casos Abertos', text='count', color='Faixa', color_discrete_map={'0 a 3 Dias':'#0056b3', '4 a 7 Dias':'#f0ad4e', '+7 Dias':'#d9534f'})
            fig_age.update_traces(textposition='outside', showlegend=False)
            fig_age.update_layout(height=300, margin=dict(l=0, r=0, t=40, b=0), plot_bgcolor='rgba(0,0,0,0)')
            st.plotly_chart(fig_age, use_container_width=True)
            
    with c2:
        df_sla = df_aba['SLA Macro'].value_counts().reset_index()
        if not df_sla.empty:
            fig_sla = px.pie(df_sla, names='SLA Macro', values='count', hole=0.5, title='Saúde do SLA (Total)', color='SLA Macro', color_discrete_map={'✅ No Prazo':'#00CC96', '🔴 Atrasado':'#EF553B'})
            fig_sla.update_layout(height=300, margin=dict(l=0, r=0, t=40, b=0), plot_bgcolor='rgba(0,0,0,0)')
            st.plotly_chart(fig_sla, use_container_width=True)

    st.markdown("---")
    st.markdown(f"### 🛠️ Tabela de Casos ({tipo_aba})")
    
    # Define as colunas específicas para cada aba
    if tipo_aba == "OA":
        colunas_ordem_ideal = [
            'Número', 'Link Salesforce', 'Conta', 'Conta - CNPJ', 'Abertura', 'Fechamento', 'Origem', 'Tipo Solicitação', 
            'Motivo', 'Detalhe', 'Substatus', 'SLA (Prazo)', 'Status', 'BaseCorp Carteira', 'Item de Contrato', 
            'Descrição', 'Fila Principal', 'Subfila', 'Idade (Dias)', 'ID do Caso', 'ID do Proprietário'
        ]
    else:
        colunas_ordem_ideal = [
            'Número', 'Link Salesforce', 'Conta', 'Conta - CNPJ', 'Conta - Posição Fin.', 'Conta - Classificação',
            'OS - Número', 'OS - Franquia', 'OS - Tipo Serviço', 'OS - Agendamento', 'OS - Técnico',
            'Item de Contrato', 'Asset - Status', 'Asset - Endereço', 'Asset - Instalação',
            'Abertura', 'Fechamento', 'Origem', 'Tipo Solicitação', 'Motivo', 'Detalhe', 'Substatus', 
            'SLA (Prazo)', 'Status', 'BaseCorp Carteira', 'Descrição', 'Fila Principal', 'Subfila', 'Idade (Dias)', 'ID do Caso', 'ID do Proprietário'
        ]
        
    df_render = df_aba[colunas_ordem_ideal].copy()
    df_render.insert(0, 'Selecionar', False)
    
    def colorir_linha(row):
        return ['background-color: #ffebee' if 'Atrasado' in row['SLA (Prazo)'] else 'background-color: #ffffff' for _ in row]

    colunas_bloqueadas = df_render.columns.drop(['Selecionar']).tolist()

    edited_df = st.data_editor(
        df_render.style.apply(colorir_linha, axis=1),
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
        key=f"editor_{fila_atual}_{tipo_aba}"
    )

    casos_selecionados = edited_df[edited_df['Selecionar'] == True]
    if not casos_selecionados.empty:
        st.markdown(f"**⚡ Ações Disponíveis ({len(casos_selecionados)} selecionados):**")
        c_btn1, c_btn2, c_btn3 = st.columns(3)
        with c_btn1:
            if st.button(f"🔄 Transferir Casos ({tipo_aba})", use_container_width=True, key=f"btn_transf_{tipo_aba}"):
                modal_transferir_comentar(casos_selecionados, lista_proprietarios)
        with c_btn2:
            if st.button(f"📝 Editar Casos ({tipo_aba})", use_container_width=True, key=f"btn_edit_{tipo_aba}"):
                modal_editar_casos(casos_selecionados, df_aba)
        with c_btn3:
            if st.button(f"🔔 Criar Follow-up ({tipo_aba})", use_container_width=True, key=f"btn_fup_{tipo_aba}"):
                modal_followup(casos_selecionados, lista_proprietarios)

    st.markdown("<br>", unsafe_allow_html=True)
    def to_excel(df_export):
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_temp = df_export.drop(columns=['Selecionar', 'ID do Proprietário', 'ID do Caso']).copy() 
            df_temp['Abertura'] = df_temp['Abertura'].dt.tz_localize(None)
            if df_temp['Fechamento'].notna().any():
                df_temp['Fechamento'] = df_temp['Fechamento'].dt.tz_localize(None)
            df_temp.to_excel(writer, index=False, sheet_name=f'Extrato_{tipo_aba}')
        return output.getvalue()

    st.download_button(
        label=f"📥 Baixar Extrato {tipo_aba} ({len(df_render)} registros)",
        data=to_excel(df_render),
        file_name=f'extrato_{tipo_aba}_{fila_atual.replace(" ", "_").lower()}.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        key=f"btn_dw_{tipo_aba}"
    )


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

st.sidebar.markdown(f"**Conexão Estabelecida:**<br> <span style='color: #0056b3; font-size: 14px;'>{SF_USER}</span>", unsafe_allow_html=True)
st.sidebar.caption(f"Última Sincronização: {st.session_state.last_update}")

if HAS_AUTOREFRESH:
    modo_tv = st.sidebar.toggle("⏱️ Atualização Automática (5 min)")
    if modo_tv:
        st_autorefresh(interval=5 * 60 * 1000, key="data_refresh")

st.sidebar.markdown("---")

if st.sidebar.button("📄 Gerar Resumo Diário", type="primary", use_container_width=True):
    per_sel = st.session_state.get('last_periodo', "Últimos 30 Dias")
    dt_ini = st.session_state.get('last_dt_inicio', None)
    dt_fim = st.session_state.get('last_dt_fim', None)
    modal_resumo_diario(st.session_state.get('last_df', pd.DataFrame()), per_sel, dt_ini, dt_fim)

busca_global = st.sidebar.text_input("🔍 Busca Rápida (Nº do Caso ou Conta)")

st.markdown("""
    <style>
    [data-testid="stSidebar"] div.stButton > button { width: 95% !important; border-radius: 6px !important; margin-top: 10px !important; border: none !important; }
    </style>
""", unsafe_allow_html=True)

if st.sidebar.button("🔄 Sincronizar Agora", use_container_width=True):
    st.cache_data.clear()
    st.session_state.last_update = datetime.now(fuso_br).strftime("%d/%m/%Y %H:%M")
    st.rerun()

if st.sidebar.button("🚪 Sair do Painel", use_container_width=True):
    st.session_state.app_authenticated = False
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

st.session_state.last_periodo = periodo_selecionado
st.session_state.last_dt_inicio = dt_inicio
st.session_state.last_dt_fim = dt_fim

filtro_meus_casos = st.sidebar.toggle("🙋‍♂️ Ver apenas Meus Casos", value=False)
incluir_fechados = st.sidebar.checkbox("Mostrar Casos Fechados", value=False)

df_filtrado = get_data(periodo_selecionado, dt_inicio, dt_fim, incluir_fechados, SF_USER, SF_PWD, SF_TOKEN)
lista_proprietarios = get_owner_options(SF_USER, SF_PWD, SF_TOKEN)
api_user_id = get_api_user_id(SF_USER, SF_PWD, SF_TOKEN)

if filtro_meus_casos and not df_filtrado.empty and api_user_id:
    df_filtrado = df_filtrado[df_filtrado['ID do Proprietário'] == api_user_id]

if busca_global and not df_filtrado.empty:
    mask = (
        df_filtrado['Número'].astype(str).str.contains(busca_global, case=False, na=False) |
        df_filtrado['Conta'].astype(str).str.contains(busca_global, case=False, na=False)
    )
    df_filtrado = df_filtrado[mask]

st.session_state.last_df = df_filtrado.copy() if not df_filtrado.empty else pd.DataFrame()

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
            
    tab_oa, tab_os = st.tabs(["📁 Visão OA (Ordem de Atendimento)", "🔧 Visão OS (Ordem de Serviço)"])

    with tab_oa:
        render_aba_conteudo(df_view[df_view['Tipo Aba'] == 'OA'], "OA", fila_atual, lista_proprietarios)

    with tab_os:
        render_aba_conteudo(df_view[df_view['Tipo Aba'] == 'OS'], "OS", fila_atual, lista_proprietarios)
