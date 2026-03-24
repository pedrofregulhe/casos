import streamlit as st
import pandas as pd
from simple_salesforce import Salesforce
from io import BytesIO
from datetime import datetime
import plotly.express as px

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Gestão de Casos OA", layout="wide", initial_sidebar_state="expanded")

# --- CSS CUSTOMIZADO (Visual Executivo e Clean) ---
st.markdown("""
    <style>
    div[data-testid="metric-container"] {
        background-color: #f8f9fa;
        border: 1px solid #e0e0e0;
        padding: 18px;
        border-radius: 10px;
        box-shadow: 2px 2px 6px rgba(0,0,0,0.04);
    }
    div[data-testid="stMetricValue"] {
        color: #0c1c2b;
        font-weight: 700;
    }
    div[data-testid="stMetricLabel"] {
        color: #6c757d;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    [data-testid="metric-container"]:nth-child(4) div[data-testid="stMetricValue"] {
        color: #d9534f;
    }
    .stButton>button {
        width: 100%;
        border-radius: 5px;
        border: 1px solid #dcdcdc;
        background-color: #ffffff;
        color: #6c757d;
        font-size: 14px;
        margin-top: 10px;
    }
    .stButton>button:hover {
        border-color: #0c1c2b;
        color: #0c1c2b;
        background-color: #fcfcfc;
    }
    h1 {
        font-size: 32px !important;
        color: #0c1c2b;
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

# --- FUNÇÃO DE BUSCA OTIMIZADA (Filtra direto no Banco de Dados) ---
# A função agora recebe os filtros como parâmetros. Se o usuário mudar o filtro, ela roda de novo.
@st.cache_data(ttl=1800) 
def get_data(periodo_selecionado, incluir_fechados):
    
    # 1. Traduzindo a escolha da tela para a linguagem do Salesforce (SOQL)
    mapa_periodos = {
        "Últimos 30 Dias": "LAST_N_DAYS:30",
        "Últimos 60 Dias": "LAST_N_DAYS:60",
        "Últimos 90 Dias": "LAST_N_DAYS:90",
        "Este Ano": "THIS_YEAR"
    }
    filtro_data = mapa_periodos[periodo_selecionado]
    
    # 2. Traduzindo o botão de incluir fechados
    filtro_status = ""
    if not incluir_fechados:
        filtro_status = "AND Status != 'Closed' AND Status != 'Fechado'"

    # 3. Construindo a Query Dinâmica
    query = f"""
    SELECT 
        Id, CaseNumber, CreatedDate, Status,
        Account.Name, Account.FOZ_CPF__c,
        Origin, Type, FOZ_Motivo__c, FOZ_Detalhe__c, Owner.Name, 
        (SELECT IsViolated FROM CaseMilestones)
    FROM Case 
    WHERE Type = 'OA'
      AND CreatedDate = {filtro_data}
      {filtro_status}
    """
    
    result = sf.query_all(query)
    sf_base_url = "https://seusalesforce.lightning.force.com/lightning/r/Case/"
    
    linhas = []
    for record in result['records']:
        dono_upper = record['Owner']['Name'].upper() if record['Owner'] else 'SISTEMA/SEM DONO'
        filas_conhecidas = ["ERRO SISTÊMICO", "CAPACIDADE", "FRANQUIAS", "AUDITORIA", "HELP TEC", "JURÍDICO", "INFORMAÇÃO", "RAF"]
        
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
            'Status': record['Status'],
            'Macro Status': macro_status,
            'SLA Atrasado': 'Sim' if sla_atrasado else 'Não',
            'Conta': record['Account']['Name'] if record['Account'] else '-',
            'Motivo': record['FOZ_Motivo__c']
        })
        
    return pd.DataFrame(linhas)

# --- BARRA LATERAL (SIDEBAR) CONTROLES DE BUSCA ---
st.sidebar.title("Filtros de Busca")
st.sidebar.markdown(f"Última Atualização: `{st.session_state.last_update}`")

if st.sidebar.button("🔄 Sincronizar Agora", type="primary"):
    st.cache_data.clear()
    st.session_state.last_update = datetime.now().strftime("%d/%m/%Y %H:%M")
    st.rerun()

st.sidebar.markdown("---")
st.sidebar.markdown("**Ajuste os dados que deseja carregar:**")

# Filtros que afetam diretamente o Salesforce
periodo_selecionado = st.sidebar.selectbox(
    "Período de Abertura", 
    ["Últimos 30 Dias", "Últimos 60 Dias", "Últimos 90 Dias", "Este Ano"],
    index=0 # Deixa "Últimos 30 Dias" como padrão
)

incluir_fechados = st.sidebar.checkbox("Mostrar Casos Fechados", value=False) # Vem desmarcado por padrão

# Busca os dados usando os filtros selecionados
df_filtrado = get_data(periodo_selecionado, incluir_fechados)

# --- TELA PRINCIPAL ---
st.title("Visão de Casos OA")
st.markdown("<br>", unsafe_allow_html=True)

if st.button("⬅️ Voltar para Visão Geral", disabled=(st.session_state.fila_selecionada is None), type="secondary"):
    st.session_state.fila_selecionada = None
    st.rerun()

# --- VISÃO 1: CARDS GERAIS ---
if st.session_state.fila_selecionada is None:
    if df_filtrado.empty:
        st.info("Nenhum caso encontrado para os filtros selecionados.")
    else:
        filas_principais = df_filtrado['Fila Principal'].unique()
        cols = st.columns(3)
        
        for i, fila in enumerate(filas_principais):
            col = cols[i % 3]
            df_fila = df_filtrado[df_filtrado['Fila Principal'] == fila]
            
            vol_total = len(df_fila)
            em_tratativa = len(df_fila[df_fila['Macro Status'] == 'Em Tratativa'])
            fechados = len(df_fila[df_fila['Macro Status'] == 'Fechado'])
            atrasados = len(df_fila[(df_fila['SLA Atrasado'] == 'Sim') & (df_fila['Macro Status'] == 'Em Tratativa')])
            
            with col:
                st.markdown(f"<h3 style='font-size: 20px; color: #0c1c2b; margin-bottom: -15px;'>{fila}</h3>", unsafe_allow_html=True)
                inner_col1, inner_col2, inner_col3 = st.columns(3)
                with inner_col1:
                    st.metric("Volume", vol_total)
                with inner_col2:
                    st.metric("🟠 Ativos", em_tratativa)
                with inner_col3:
                    st.metric("🔴 Atrasados", atrasados)
                
                if st.button("Detalhar", key=f"btn_{fila}"):
                    st.session_state.fila_selecionada = fila
                    st.rerun()
            
            if (i + 1) % 3 == 0:
                st.markdown("<br>", unsafe_allow_html=True)

# --- VISÃO 2: DETALHE DA FILA ---
else:
    fila_atual = st.session_state.fila_selecionada
    st.subheader(f"Visão Detalhada: {fila_atual}")
    df_extrato = df_filtrado[df_filtrado['Fila Principal'] == fila_atual].copy()
    
    col_chart1, col_chart2 = st.columns(2)
    
    with col_chart1:
        label_grf1 = 'Casos OA por Usuário' if fila_atual != 'CORPORATIVO' else 'Casos OA por Carteira'
        df_grp = df_extrato['Subfila'].value_counts().reset_index()
        fig_grp = px.bar(df_grp, x='count', y='Subfila', orientation='h', title=label_grf1, labels={'count': 'Volume'})
        fig_grp.update_layout(height=350, yaxis={'categoryorder':'total ascending'})
        st.plotly_chart(fig_grp, use_container_width=True)
        
    with col_chart2:
        df_sla = df_extrato['SLA Atrasado'].value_counts().reset_index()
        fig_sla = px.pie(df_sla, names='SLA Atrasado', values='count', hole=0.5, title='Distribuição do SLA (Total)', color='SLA Atrasado',
                         color_discrete_map={'Não':'#00CC96', 'Sim':'#EF553B'})
        fig_sla.update_layout(height=350)
        st.plotly_chart(fig_sla, use_container_width=True)

    st.markdown("---")
    def to_excel(df_export):
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_export.to_excel(writer, index=False, sheet_name='ExtratoOA')
        return output.getvalue()

    st.download_button(
        label="📥 Baixar Extrato da Fila em Excel",
        data=to_excel(df_extrato),
        file_name=f'extrato_{fila_atual.replace(" ", "_").lower()}.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        type="primary"
    )
    
    st.dataframe(
        df_extrato,
        column_config={
            "Link Salesforce": st.column_config.LinkColumn("Acessar", display_text="Abrir"),
            "Abertura": st.column_config.DateColumn("Abertura", format="DD/MM/YYYY"),
            "ID do Caso": None
        },
        use_container_width=True,
        hide_index=True
    )
