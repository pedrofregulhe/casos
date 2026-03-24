import streamlit as st
import pandas as pd
from simple_salesforce import Salesforce
from io import BytesIO
from datetime import datetime

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Gestão de Casos OA", layout="wide", initial_sidebar_state="expanded")

# CSS customizado para deixar a interface mais elegante e os cards mais bonitos
st.markdown("""
    <style>
    div[data-testid="metric-container"] {
        background-color: #f8f9fa;
        border: 1px solid #e0e0e0;
        padding: 15px;
        border-radius: 8px;
        box-shadow: 2px 2px 5px rgba(0,0,0,0.05);
    }
    .stButton>button {
        width: 100%;
        border-radius: 5px;
        font-weight: bold;
    }
    </style>
""", unsafe_allow_html=True)

# --- CONTROLE DE ESTADO (Para o clique nos cards) ---
if 'fila_selecionada' not in st.session_state:
    st.session_state.fila_selecionada = None

# --- CONEXÃO E DADOS ---
@st.cache_resource
def init_connection():
    return Salesforce(
        username=st.secrets["sf_username"],
        password=st.secrets["sf_password"],
        security_token=st.secrets["sf_token"],
        domain='login'
    )

sf = init_connection()

# O botão de sincronizar limpa o cache dessa função
@st.cache_data(ttl=3600)
def get_data():
    # Adicionei o Id (para o link) e CreatedDate (para o filtro)
    query = """
    SELECT 
        Id, CaseNumber, CreatedDate, Status,
        Account.Name, Account.FOZ_StatusPosicaoFinanceira__c, Account.FOZ_CPF__c,
        Origin, Type, FOZ_TipoSolicitacao__c, FOZ_Motivo__c, FOZ_Detalhe__c, FOZ_Subdetalhe__c, 
        Owner.Name, 
        (SELECT IsViolated FROM CaseMilestones)
    FROM Case 
    WHERE Type = 'OA'
    AND CreatedDate = LAST_N_DAYS:90
    """
    result = sf.query_all(query)
    
    # URL base do seu Salesforce (ajuste se necessário)
    sf_base_url = "https://seusalesforce.lightning.force.com/lightning/r/Case/"
    
    linhas = []
    for record in result['records']:
        dono_nome = record['Owner']['Name'] if record['Owner'] else 'SISTEMA/SEM DONO'
        dono_upper = dono_nome.upper()
        
        # --- LÓGICA DE CATEGORIZAÇÃO DAS FILAS ---
        filas_conhecidas = ["ERRO SISTÊMICO", "CAPACIDADE", "FRANQUIAS", "AUDITORIA", 
                            "HELP TEC", "JURÍDICO", "INFORMAÇÃO", "RAF"]
        
        if dono_upper in filas_conhecidas:
            fila_principal = dono_upper
            subfila = "-"
        elif dono_upper.startswith("CARTEIRA"):
            fila_principal = "CORPORATIVO"
            subfila = dono_upper
        else:
            fila_principal = "ATRIBUÍDO AO USUÁRIO"
            subfila = dono_upper # Guarda quem é o usuário
            
        # --- STATUS E SLA ---
        status_atual = record['Status']
        macro_status = "Fechado" if status_atual in ['Closed', 'Fechado'] else "Em Tratativa"
        
        sla_atrasado = False
        if record['CaseMilestones'] and record['CaseMilestones']['records']:
            for milestone in record['CaseMilestones']['records']:
                if milestone['IsViolated']:
                    sla_atrasado = True
                    break
                    
        linhas.append({
            'ID do Caso': record['Id'],
            'Link Salesforce': f"{sf_base_url}{record['Id']}/view",
            'Número': record['CaseNumber'],
            'Data de Abertura': pd.to_datetime(record['CreatedDate']).tz_localize(None),
            'Fila Principal': fila_principal,
            'Subfila': subfila,
            'Status': status_atual,
            'Macro Status': macro_status,
            'SLA Atrasado': 'Sim' if sla_atrasado else 'Não',
            'Conta': record['Account']['Name'] if record['Account'] else '-',
            'Motivo': record['FOZ_Motivo__c']
        })
        
    return pd.DataFrame(linhas)

# --- BARRA LATERAL (SIDEBAR) ---
st.sidebar.title("⚙️ Filtros e Controles")

if st.sidebar.button("🔄 Sincronizar com Salesforce", type="primary"):
    st.cache_data.clear()
    st.rerun()

df = get_data()

st.sidebar.markdown("---")
# Filtro de Data
min_date = df['Data de Abertura'].min().date() if not df.empty else datetime.today().date()
max_date = df['Data de Abertura'].max().date() if not df.empty else datetime.today().date()
datas_selecionadas = st.sidebar.date_input("Período de Abertura", [min_date, max_date], min_value=min_date, max_value=max_date)

# Filtro de Status
lista_status = df['Status'].unique().tolist()
status_selecionados = st.sidebar.multiselect("Status do Caso", lista_status, default=lista_status)

# Aplicando os filtros no DataFrame
if len(datas_selecionadas) == 2:
    data_inicio, data_fim = datas_selecionadas
    df_filtrado = df[
        (df['Data de Abertura'].dt.date >= data_inicio) & 
        (df['Data de Abertura'].dt.date <= data_fim) &
        (df['Status'].isin(status_selecionados))
    ]
else:
    df_filtrado = df[df['Status'].isin(status_selecionados)]

# --- TELA PRINCIPAL: CARDS ---
st.title("🏛️ Visão Executiva de Casos (OA)")

if st.button("⬅️ Voltar para Visão Geral", disabled=(st.session_state.fila_selecionada is None)):
    st.session_state.fila_selecionada = None
    st.rerun()

# Se não houver fila selecionada, mostra os cards gerais
if st.session_state.fila_selecionada is None:
    st.subheader("Resumo por Fila")
    
    filas_principais = df_filtrado['Fila Principal'].unique()
    
    # Criando as linhas de cards dinamicamente (3 por linha)
    cols = st.columns(3)
    for i, fila in enumerate(filas_principais):
        col = cols[i % 3]
        df_fila = df_filtrado[df_filtrado['Fila Principal'] == fila]
        
        vol_total = len(df_fila)
        em_tratativa = len(df_fila[df_fila['Macro Status'] == 'Em Tratativa'])
        fechados = len(df_fila[df_fila['Macro Status'] == 'Fechado'])
        atrasados = len(df_fila[(df_fila['SLA Atrasado'] == 'Sim') & (df_fila['Macro Status'] == 'Em Tratativa')])
        
        with col:
            with st.container():
                st.markdown(f"### {fila}")
                st.markdown(f"**Volume Total:** {vol_total}")
                st.markdown(f"🟢 **Fechados:** {fechados} | 🟡 **Em Tratativa:** {em_tratativa}")
                st.markdown(f"🔴 **SLA Atrasado (Ativos):** {atrasados}")
                
                # O botão que simula o clique no card
                if st.button(f"🔍 Detalhar {fila}", key=f"btn_{fila}"):
                    st.session_state.fila_selecionada = fila
                    st.rerun()
            st.markdown("<br>", unsafe_allow_html=True) # Espaçamento

# --- TELA SECUNDÁRIA: EXTRATO DETALHADO ---
else:
    fila_atual = st.session_state.fila_selecionada
    st.subheader(f"📑 Extrato Detalhado: {fila_atual}")
    
    df_extrato = df_filtrado[df_filtrado['Fila Principal'] == fila_atual].copy()
    
    # Opção para exportar para Excel
    def to_excel(df_export):
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_export.to_excel(writer, index=False, sheet_name='Extrato')
        processed_data = output.getvalue()
        return processed_data

    excel_data = to_excel(df_extrato)
    st.download_button(
        label="📥 Baixar Extrato em Excel",
        data=excel_data,
        file_name=f'extrato_{fila_atual.replace(" ", "_").lower()}.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        type="primary"
    )
    
    # Renderizando a tabela com link clicável nativo do Streamlit
    st.dataframe(
        df_extrato,
        column_config={
            "Link Salesforce": st.column_config.LinkColumn(
                "Acessar Caso", display_text="Abrir no Salesforce"
            ),
            "ID do Caso": None # Esconde o ID feio da visualização
        },
        use_container_width=True,
        hide_index=True
    )
