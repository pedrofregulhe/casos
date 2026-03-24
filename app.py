import streamlit as st
import pandas as pd
from simple_salesforce import Salesforce
from io import BytesIO
from datetime import datetime
import plotly.express as px

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Gestão de Casos OA", layout="wide", initial_sidebar_state="expanded")

# --- CSS CUSTOMIZADO (REPLICAÇÃO EXATA DA REFERÊNCIA) ---
# Aqui construímos o design dos cards, sombras, bordas superiores azuis e tipografia.
st.markdown("""
    <style>
    /* Estilo do Título Executivo */
    h1 {
        font-size: 26px !important;
        font-family: 'IBM Plex Sans', sans-serif !important;
        color: #1c2b39;
        margin-bottom: -15px !important;
    }
    h2, h3 {
        font-family: 'IBM Plex Sans', sans-serif !important;
        color: #1c2b39;
    }

    /* Container para alinhar os cards horizontalmente */
    .kpi-row-container {
        display: flex;
        flex-direction: row;
        justify-content: space-around;
        gap: 10px;
        width: 100%;
        margin-bottom: 20px;
    }

    /* O Card de KPI Principal - Baseado na imagem de referência */
    .kpi-card {
        background-color: #ffffff;
        border-radius: 8px;
        box-shadow: 0 4px 10px rgba(0,0,0,0.06); /* Sombra sutil */
        border-top: 6px solid #0056b3; /* Borda superior azul */
        padding: 20px;
        flex: 1; /* Todos os cards com o mesmo tamanho */
        text-align: center;
        min-width: 160px;
    }
    
    /* Card Especial para Alerta de SLA Atrasado (Vermelho) */
    .kpi-card-alert {
        background-color: #ffffff;
        border-radius: 8px;
        box-shadow: 0 4px 10px rgba(0,0,0,0.06);
        border-top: 6px solid #d9534f; /* Borda superior vermelha */
        padding: 20px;
        flex: 1;
        text-align: center;
        min-width: 160px;
    }

    /* Rótulo superior do Card (Texto Cinza, Centrado, Uppercase) */
    .kpi-label {
        font-family: 'IBM Plex Sans', sans-serif;
        color: #727b84;
        font-size: 13px;
        text-transform: uppercase;
        letter-spacing: 0.8px;
        margin-bottom: 8px;
    }

    /* Valor principal do Card (Texto Azul Grande, Bold) */
    .kpi-value {
        font-family: 'IBM Plex Sans', sans-serif;
        color: #0056b3;
        font-size: 34px;
        font-weight: 700;
        margin-bottom: 0px;
    }
    
    /* Valor principal Especial para Alerta (Vermelho) */
    .kpi-value-alert {
        font-family: 'IBM Plex Sans', sans-serif;
        color: #d9534f;
        font-size: 34px;
        font-weight: 700;
        margin-bottom: 0px;
    }
    
    /* Rótulo da Fila Acima dos Cards */
    .queue-title {
        margin-top: 30px;
        margin-bottom: 5px;
        border-bottom: 2px solid #eaeaea;
        padding-bottom: 5px;
    }

    /* Botão Detalhar discreto abaixo dos cards */
    .stButton>button {
        width: 100%;
        border-radius: 6px;
        border: 1px solid #dcdcdc;
        background-color: #ffffff;
        color: #6c757d;
        font-size: 14px;
        margin-top: 5px;
    }
    .stButton>button:hover {
        border-color: #0c1c2b;
        color: #0c1c2b;
        background-color: #fcfcfc;
    }
    
    /* Sidebar visual executivo */
    .st-emotion-cache-1vt4y43 {
        color: #0c1c2b !important;
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

# --- FUNÇÃO DE BUSCA OTIMIZADA (Conversa com o BD) ---
@st.cache_data(ttl=1800) 
def get_data(periodo_selecionado, incluir_fechados):
    mapa_periodos = {
        "Últimos 30 Dias": "LAST_N_DAYS:30",
        "Últimos 60 Dias": "LAST_N_DAYS:60",
        "Últimos 90 Dias": "LAST_N_DAYS:90",
        "Este Ano": "THIS_YEAR"
    }
    filtro_data = mapa_periodos[periodo_selecionado]
    filtro_status = ""
    if not incluir_fechados:
        filtro_status = "AND Status != 'Closed' AND Status != 'Fechado'"

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
    # URL BASE CORRIGIDA
    sf_base_url = "https://ibbl.lightning.force.com/lightning/r/Case/"
    
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

# --- FUNÇÃO AUXILIAR PARA GERAR O HTML DO CARD (O PONTO CHAVE) ---
def render_kpi_card(label, value, alert=False):
    """Gera o HTML para um card KPI único seguindo a imagem de referência."""
    card_class = "kpi-card" if not alert else "kpi-card-alert"
    value_class = "kpi-value" if not alert else "kpi-value-alert"
    
    html = f"""
    <div class="{card_class}">
        <div class="kpi-label">{label}</div>
        <div class="{value_class}">{value}</div>
    </div>
    """
    return html

# --- BARRA LATERAL (SIDEBAR) CONTROLES DE BUSCA ---
st.sidebar.title("Filtros de Busca")
st.sidebar.markdown(f"Última Atualização: `{st.session_state.last_update}`")

if st.sidebar.button("🔄 Sincronizar Agora", type="primary"):
    st.cache_data.clear()
    st.session_state.last_update = datetime.now().strftime("%d/%m/%Y %H:%M")
    st.rerun()

st.sidebar.markdown("---")
st.sidebar.markdown("**Ajuste os dados que deseja carregar:**")

periodo_selecionado = st.sidebar.selectbox(
    "Período de Abertura", 
    ["Últimos 30 Dias", "Últimos 60 Dias", "Últimos 90 Dias", "Este Ano"],
    index=0
)
incluir_fechados = st.sidebar.checkbox("Mostrar Casos Fechados", value=False)

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
        st.subheader("Resumo Executivo por Fila")
        
        filas_principais = df_filtrado['Fila Principal'].unique()
        
        for fila in filas_principais:
            df_fila = df_filtrado[df_filtrado['Fila Principal'] == fila]
            
            vol_total = len(df_fila)
            em_tratativa = len(df_fila[df_fila['Macro Status'] == 'Em Tratativa'])
            fechados = len(df_fila[df_fila['Macro Status'] == 'Fechado'])
            atrasados = len(df_fila[(df_fila['SLA Atrasado'] == 'Sim') & (df_fila['Macro Status'] == 'Em Tratativa')])
            
            # Cabeçalho da Fila
            st.markdown(f"<h3 class='queue-title'>{fila}</h3>", unsafe_allow_html=True)
            
            # --- CONSTRUÇÃO DA LINHA DE CARDS USANDO HTML ---
            # Aqui juntamos os 4 cards numa mesma linha horizontal.
            cards_html = f"""
            <div class="kpi-row-container">
                {render_kpi_card("Volume Total", vol_total)}
                {render_kpi_card("Em Tratativa", em_tratativa)}
                {render_kpi_card("Fechados", fechados)}
                {render_kpi_card("SLA Atrasado (Ativos)", atrasados, alert=(atrasados > 0))}
            </div>
            """
            st.markdown(cards_html, unsafe_allow_html=True)
            
            # Botão Detalhar discreto abaixo dos cards (mantendo funcionalidade Streamlit)
            if st.button("🔍 Detalhar", key=f"btn_{fila}"):
                st.session_state.fila_selecionada = fila
                st.rerun()
            
            st.markdown("<br>", unsafe_allow_html=True) # Espaçamento

# --- VISÃO 2: DETALHE DA FILA (Inalterada, pois já estava executiva) ---
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
            # CONFIGURAÇÃO CORRIGIDA PARA ESCONDER O ID FEIO
            "ID do Caso": None
        },
        use_container_width=True,
        hide_index=True
    )
