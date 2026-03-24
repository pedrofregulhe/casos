import streamlit as st
import pandas as pd
from simple_salesforce import Salesforce
from io import BytesIO
from datetime import datetime
import plotly.express as px

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Gestão de Casos", layout="wide", initial_sidebar_state="expanded")

# --- CSS CUSTOMIZADO (CARDS QUADRADOS E LIMPEZA) ---
st.markdown("""
    <style>
    .stApp { background-color: #f8f9fa !important; }
    header { visibility: hidden !important; height: 0px !important; display: none !important; }
    #MainMenu { visibility: hidden !important; display: none !important; }
    footer { visibility: hidden !important; display: none !important; }
    .block-container { padding-top: 1.5rem !important; padding-bottom: 1rem !important; }
    
    h1 { font-size: 24px !important; margin-bottom: 20px !important; color: #1a2935; }
    
    /* Remove a margem superior dos botões que ficam embaixo dos cards */
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
    </style>
""", unsafe_allow_html=True)

# --- CONTROLE DE ESTADO ---
if 'fila_selecionada' not in st.session_state:
    st.session_state.fila_selecionada = None
if 'last_update' not in st.session_state:
    st.session_state.last_update = datetime.now().strftime("%d/%m/%Y %H:%M")

# --- CONEXÃO COM SALESFORCE ---
@st.cache_resource
def init_connection():
    return Salesforce(
        username=st.secrets["sf_username"],
        password=st.secrets["sf_password"],
        security_token=st.secrets["sf_token"],
        domain='login'
    )

sf = init_connection()

# --- FUNÇÃO DE BUSCA OTIMIZADA ---
@st.cache_data(ttl=1800) 
def get_data(periodo_selecionado, dt_inicio, dt_fim, incluir_fechados):
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

    filtro_status = ""
    if not incluir_fechados:
        filtro_status = "AND Status != 'Closed' AND Status != 'Fechado'"

    query = f"""
    SELECT 
        Id, CaseNumber, CreatedDate, Status,
        Account.Name, Account.FOZ_CPF__c,
        Origin, Type, FOZ_TipoSolicitacao__c, FOZ_Motivo__c, FOZ_Detalhe__c, FOZ_Subdetalhe__c, Owner.Name, 
        (SELECT IsViolated FROM CaseMilestones)
    FROM Case 
    WHERE (Type = 'OA' OR (Owner.Name LIKE 'CARTEIRA%' AND (Type != 'OS' OR Type = null)))
      AND {filtro_data}
      {filtro_status}
    """
    result = sf.query_all(query)
    sf_base_url = "https://ibbl.lightning.force.com/lightning/r/Case/"
    
    linhas = []
    for record in result['records']:
        dono_upper = record['Owner']['Name'].upper() if record['Owner'] else 'SISTEMA/SEM DONO'
        
        filas_conhecidas = [
            "ERRO SISTÊMICO", "CAPACIDADE", "FRANQUIAS", "AUDITORIA", 
            "HELP TEC", "JURÍDICO", "INFORMAÇÃO", "RAF", "FINANCEIRO", "BACKOFFICE"
        ]
        
        if dono_upper in filas_conhecidas:
            fila_principal = dono_upper
            subfila = "-"
        elif dono_upper.startswith("CARTEIRA"):
            fila_principal = "CORPORATIVO"
            subfila = dono_upper
        else:
            fila_principal = "ATRIBUÍDO AO USUÁRIO"
            subfila = dono_upper
            
        macro_status = "Fechado" if record['Status'] in ['Closed', 'Fechado'] else "Em Tratativa"
        sla_atrasado = any(m['IsViolated'] for m in record['CaseMilestones']['records']) if record['CaseMilestones'] else False
                    
        linhas.append({
            'ID do Caso': record['Id'],
            'Link Salesforce': f"{sf_base_url}{record['Id']}/view",
            'Número': record['CaseNumber'],
            'Abertura': pd.to_datetime(record['CreatedDate']).tz_localize(None),
            'Fila Principal': fila_principal,
            'Subfila': subfila,
            'Origem': record['Origin'],
            'Tipo Salesforce': record['Type'] if record['Type'] else 'Sem Tipo',
            'Tipo Solicitação': record['FOZ_TipoSolicitacao__c'],
            'Motivo': record['FOZ_Motivo__c'],
            'Detalhe': record['FOZ_Detalhe__c'],
            'Subdetalhe': record['FOZ_Subdetalhe__c'],
            'Status': record['Status'],
            'Macro Status': macro_status,
            'SLA Atrasado': 'Sim' if sla_atrasado else 'Não',
            'Conta': record['Account']['Name'] if record['Account'] else '-'
        })
        
    return pd.DataFrame(linhas)

# --- FUNÇÃO PARA DESENHAR O CARD QUADRADO ---
def desenhar_card(fila_nome, df_fila):
    vol = len(df_fila)
    trat = len(df_fila[df_fila['Macro Status'] == 'Em Tratativa'])
    fech = len(df_fila[df_fila['Macro Status'] == 'Fechado'])
    atr = len(df_fila[(df_fila['SLA Atrasado'] == 'Sim') & (df_fila['Macro Status'] == 'Em Tratativa')])
    
    cor_atraso = "#d9534f" if atr > 0 else "#555555"
    
    # HTML do Card Visual Quadrado
    html_card = f"""
    <div style="background-color: white; border: 1px solid #dce1e6; border-radius: 8px 8px 0px 0px; padding: 15px; height: 145px; box-shadow: 0 2px 4px rgba(0,0,0,0.02); display: flex; flex-direction: column; justify-content: space-between;">
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
    # Botão que fica "grudado" no card HTML via CSS
    if st.button(f"Abrir Detalhe", key=f"btn_{fila_nome}", use_container_width=True):
        st.session_state.fila_selecionada = fila_nome
        st.rerun()

# --- MENU LATERAL (SIDEBAR) ---
st.sidebar.title("Filtros")
st.sidebar.caption(f"Última Sincronização: {st.session_state.last_update}")

if st.sidebar.button("🔄 Sincronizar Agora", type="primary"):
    st.cache_data.clear()
    st.session_state.last_update = datetime.now().strftime("%d/%m/%Y %H:%M")
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
df_filtrado = get_data(periodo_selecionado, dt_inicio, dt_fim, incluir_fechados)

# --- ORDENAÇÃO DAS FILAS ---
todas_filas = df_filtrado['Fila Principal'].unique().tolist() if not df_filtrado.empty else []
if "ATRIBUÍDO AO USUÁRIO" in todas_filas:
    todas_filas.remove("ATRIBUÍDO AO USUÁRIO")
    filas_ordenadas = sorted(todas_filas)
    filas_ordenadas.append("ATRIBUÍDO AO USUÁRIO")
else:
    filas_ordenadas = sorted(todas_filas)

# --- RENDERIZAÇÃO DA TELA (VISÃO GERAL VS ABA LATERAL) ---
if df_filtrado.empty:
    st.markdown("<h1>Visão Operacional - Casos OA</h1>", unsafe_allow_html=True)
    st.info("Nenhum caso encontrado para os filtros selecionados.")
    
elif st.session_state.fila_selecionada is None:
    # VISÃO 1: GRID DE CARDS QUADRADOS
    st.markdown("<h1>Visão Operacional - Escolha uma Fila</h1>", unsafe_allow_html=True)
    
    # Cria colunas para organizar os cards em formato de grade (4 por linha)
    cols = st.columns(4)
    for i, fila in enumerate(filas_ordenadas):
        df_fila = df_filtrado[df_filtrado['Fila Principal'] == fila]
        with cols[i % 4]:
            desenhar_card(fila, df_fila)
            st.markdown("<br>", unsafe_allow_html=True)

else:
    # VISÃO 2: EFEITO "ABA LATERAL" (Master-Detail)
    # A tela divide: 25% para os cards empilhados na esquerda, 75% para os detalhes na direita
    fila_atual = st.session_state.fila_selecionada
    col_menu, col_detalhe = st.columns([1, 3], gap="large")
    
    with col_menu:
        if st.button("⬅️ Voltar para Grade", use_container_width=True, type="primary"):
            st.session_state.fila_selecionada = None
            st.rerun()
        st.markdown("<br>", unsafe_allow_html=True)
        
        # Desenha os cards em uma lista vertical fina
        for fila in filas_ordenadas:
            df_fila = df_filtrado[df_filtrado['Fila Principal'] == fila]
            desenhar_card(fila, df_fila)
            st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)
            
    with col_detalhe:
        st.markdown(f"<h2 style='margin-top: 0px; color: #0056b3;'>Detalhes da Fila: {fila_atual}</h2>", unsafe_allow_html=True)
        df_view = df_filtrado[df_filtrado['Fila Principal'] == fila_atual].copy()
        
        # Filtro de Carteira Corporativa (aparece só no detalhe se for o caso)
        if fila_atual == "CORPORATIVO":
            carteiras_disp = sorted(df_view['Subfila'].unique().tolist())
            cart_sel = st.selectbox("📌 Filtrar Carteira Específica:", ["Todas"] + carteiras_disp)
            if cart_sel != "Todas":
                df_view = df_view[df_view['Subfila'] == cart_sel]
                
        # Gráficos Executivos (Idade e Ofensores)
        c1, c2 = st.columns(2)
        with c1:
            df_abertos = df_view[df_view['Macro Status'] == 'Em Tratativa'].copy()
            if not df_abertos.empty:
                df_abertos['Idade'] = (datetime.now() - df_abertos['Abertura']).dt.days
                bins = [-1, 3, 7, 10000]
                labels = ['0 a 3 Dias', '4 a 7 Dias', '+7 Dias']
                df_abertos['Faixa'] = pd.cut(df_abertos['Idade'], bins=bins, labels=labels)
                df_age = df_abertos['Faixa'].value_counts().reindex(labels).reset_index()
                
                fig_age = px.bar(df_age, x='Faixa', y='count', title='Idade dos Casos Abertos', text='count', color='Faixa',
                                 color_discrete_map={'0 a 3 Dias':'#0056b3', '4 a 7 Dias':'#f0ad4e', '+7 Dias':'#d9534f'})
                fig_age.update_traces(textposition='outside', showlegend=False)
                fig_age.update_layout(height=260, margin=dict(l=0, r=0, t=30, b=0), plot_bgcolor='rgba(0,0,0,0)')
                st.plotly_chart(fig_age, use_container_width=True)
            else:
                st.info("Nenhum caso em aberto.")
                
        with c2:
            df_view['Motivo Real'] = df_view['Motivo'].fillna(df_view['Tipo Solicitação']).fillna('Sem Classificação')
            df_of = df_view['Motivo Real'].value_counts().reset_index().head(5).sort_values(by='count')
            fig_of = px.bar(df_of, x='count', y='Motivo Real', orientation='h', title='Top 5 Ofensores', text='count')
            fig_of.update_traces(textposition='outside', marker_color='#17a2b8')
            fig_of.update_layout(height=260, margin=dict(l=0, r=0, t=30, b=0), plot_bgcolor='rgba(0,0,0,0)')
            st.plotly_chart(fig_of, use_container_width=True)

        st.markdown("---")
        
        def to_excel(df_export):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_export.to_excel(writer, index=False, sheet_name='Extrato')
            return output.getvalue()

        st.download_button(
            label=f"📥 Baixar Extrato Completo ({len(df_view)} registros)",
            data=to_excel(df_view),
            file_name=f'extrato_{fila_atual.replace(" ", "_").lower()}.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        
        st.dataframe(
            df_view,
            column_config={
                "Link Salesforce": st.column_config.LinkColumn("Acessar", display_text="Abrir"),
                "Abertura": st.column_config.DateColumn("Abertura", format="DD/MM/YYYY"),
                "ID do Caso": None
            },
            use_container_width=True,
            hide_index=True
        )
