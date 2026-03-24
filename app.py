import streamlit as st
import pandas as pd
from simple_salesforce import Salesforce
from io import BytesIO
from datetime import datetime
import plotly.express as px

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Gestão de Casos", layout="wide", initial_sidebar_state="expanded")

# --- CSS CUSTOMIZADO (BRANCO, SEM CABEÇALHO, PASTINHAS) ---
st.markdown("""
    <style>
    /* 1. Remove Fundo Cinza e Cabeçalhos Nativos do Streamlit */
    .stApp { background-color: #ffffff !important; }
    header { visibility: hidden !important; height: 0px !important; display: none !important; }
    #MainMenu { visibility: hidden !important; display: none !important; }
    footer { visibility: hidden !important; display: none !important; }
    .block-container { padding-top: 2rem !important; } /* Sobe o conteúdo */
    
    h1, h2, h3 { color: #1a2935; font-family: 'Segoe UI', Tahoma, sans-serif; }
    
    /* 2. Estilo das Pastinhas (Expanders) */
    div[data-testid="stExpander"] {
        border: 1px solid #dce1e6 !important;
        border-radius: 8px !important;
        box-shadow: 0 2px 5px rgba(0,0,0,0.03) !important;
        margin-bottom: 12px !important;
        background-color: #ffffff !important;
    }
    div[data-testid="stExpander"] summary {
        background-color: #f8fbff !important; /* Azul bem clarinho para a aba da pasta */
        border-radius: 8px !important;
        padding: 10px 15px !important;
    }
    div[data-testid="stExpander"] summary:hover {
        background-color: #f0f7ff !important;
    }
    div[data-testid="stExpander"] summary p {
        font-size: 18px !important;
        font-weight: 600 !important;
        color: #0056b3 !important;
    }
    </style>
""", unsafe_allow_html=True)

# --- CONTROLE DE ESTADO ---
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

    # QUERY ATUALIZADA: Pega tudo que é OA *OU* tudo que está na CARTEIRA (ignorando OS)
    query = f"""
    SELECT 
        Id, CaseNumber, CreatedDate, Status,
        Account.Name, Account.FOZ_CPF__c,
        Origin, Type, FOZ_Motivo__c, FOZ_Detalhe__c, Owner.Name, 
        (SELECT IsViolated FROM CaseMilestones)
    FROM Case 
    WHERE (Type = 'OA' OR (Owner.Name LIKE 'CARTEIRA%' AND Type != 'OS'))
      AND CreatedDate = {filtro_data}
      {filtro_status}
    """
    result = sf.query_all(query)
    sf_base_url = "https://ibbl.lightning.force.com/lightning/r/Case/"
    
    linhas = []
    for record in result['records']:
        dono_upper = record['Owner']['Name'].upper() if record['Owner'] else 'SISTEMA/SEM DONO'
        
        # LISTA DE FILAS ATUALIZADA (COM BACKOFFICE)
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
            'Status': record['Status'],
            'Macro Status': macro_status,
            'SLA Atrasado': 'Sim' if sla_atrasado else 'Não',
            'Conta': record['Account']['Name'] if record['Account'] else '-',
            'Motivo': record['FOZ_Motivo__c']
        })
        
    return pd.DataFrame(linhas)

# --- GERADOR VISUAL DE CARDS ---
def render_kpi_row(metricas):
    html = '<div style="display: flex; justify-content: space-between; gap: 15px; margin-bottom: 25px; margin-top: 10px; width: 100%;">'
    for metrica in metricas:
        label = metrica['label']
        valor = metrica['valor']
        is_alert = metrica.get('alert', False)
        
        cor_borda = "#d9534f" if is_alert else "#0056b3"
        cor_texto = "#d9534f" if is_alert else "#0056b3"
        
        html += f'''
<div style="background-color: #ffffff; border: 1px solid #e0e0e0; border-top: 4px solid {cor_borda}; border-radius: 8px; padding: 20px 10px; flex: 1; text-align: center; box-shadow: 0 4px 6px rgba(0,0,0,0.02);">
<div style="color: #6a747f; font-size: 12px; font-weight: 700; text-transform: uppercase; margin-bottom: 8px;">{label}</div>
<div style="color: {cor_texto}; font-size: 32px; font-weight: 700;">{valor}</div>
</div>
'''
    html += '</div>'
    return html

# --- BARRA LATERAL (SIDEBAR) ---
st.sidebar.title("Filtros")
st.sidebar.caption(f"Última Sincronização: {st.session_state.last_update}")

if st.sidebar.button("🔄 Sincronizar Agora", type="primary"):
    st.cache_data.clear()
    st.session_state.last_update = datetime.now().strftime("%d/%m/%Y %H:%M")
    st.rerun()

st.sidebar.markdown("---")
periodo_selecionado = st.sidebar.selectbox("Período de Abertura", ["Últimos 30 Dias", "Últimos 60 Dias", "Últimos 90 Dias", "Este Ano"], index=0)
incluir_fechados = st.sidebar.checkbox("Mostrar Casos Fechados", value=False)

df_filtrado = get_data(periodo_selecionado, incluir_fechados)

# --- TELA PRINCIPAL ---
st.markdown("<h1 style='font-size: 28px; margin-bottom: 20px;'>Visão de Casos (OA e Carteiras)</h1>", unsafe_allow_html=True)

if df_filtrado.empty:
    st.info("Nenhum caso encontrado para os filtros selecionados.")
else:
    # ORDENAÇÃO DAS FILAS (Garante que "ATRIBUÍDO AO USUÁRIO" seja sempre a última)
    todas_filas = df_filtrado['Fila Principal'].unique().tolist()
    
    if "ATRIBUÍDO AO USUÁRIO" in todas_filas:
        todas_filas.remove("ATRIBUÍDO AO USUÁRIO")
        filas_ordenadas = sorted(todas_filas)
        filas_ordenadas.append("ATRIBUÍDO AO USUÁRIO")
    else:
        filas_ordenadas = sorted(todas_filas)
    
    # --- CRIAÇÃO DAS PASTINHAS (EXPANDERS) ---
    for fila in filas_ordenadas:
        df_fila = df_filtrado[df_filtrado['Fila Principal'] == fila]
        
        vol_total = len(df_fila)
        em_tratativa = len(df_fila[df_fila['Macro Status'] == 'Em Tratativa'])
        fechados = len(df_fila[df_fila['Macro Status'] == 'Fechado'])
        atrasados = len(df_fila[(df_fila['SLA Atrasado'] == 'Sim') & (df_fila['Macro Status'] == 'Em Tratativa')])
        
        # Cria a pastinha com o nome da fila e o volume total para bater o olho rápido
        with st.expander(f"📁 {fila} ({vol_total} Casos)", expanded=False):
            
            # 1. Desenha os Cards dentro da pasta
            metricas = [
                {'label': 'Volume Total', 'valor': vol_total},
                {'label': 'Em Tratativa', 'valor': em_tratativa},
                {'label': 'Fechados', 'valor': fechados},
                {'label': 'SLA Atrasado (Ativos)', 'valor': atrasados, 'alert': atrasados > 0}
            ]
            st.markdown(render_kpi_row(metricas), unsafe_allow_html=True)
            
            # 2. Desenha os Gráficos dentro da pasta
            col_chart1, col_chart2 = st.columns(2)
            with col_chart1:
                label_grf1 = 'Casos por Usuário' if fila != 'CORPORATIVO' else 'Casos por Carteira'
                df_grp = df_fila['Subfila'].value_counts().reset_index()
                fig_grp = px.bar(df_grp, x='count', y='Subfila', orientation='h', title=label_grf1, labels={'count': 'Volume', 'Subfila': ''})
                fig_grp.update_layout(height=280, margin=dict(l=0, r=0, t=40, b=0), yaxis={'categoryorder':'total ascending'}, plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)')
                st.plotly_chart(fig_grp, use_container_width=True)
                
            with col_chart2:
                df_sla = df_fila['SLA Atrasado'].value_counts().reset_index()
                fig_sla = px.pie(df_sla, names='SLA Atrasado', values='count', hole=0.6, title='Saúde do SLA (Total)', color='SLA Atrasado', color_discrete_map={'Não':'#0056b3', 'Sim':'#d9534f'})
                fig_sla.update_layout(height=280, margin=dict(l=0, r=0, t=40, b=0), paper_bgcolor='rgba(0,0,0,0)')
                st.plotly_chart(fig_sla, use_container_width=True)

            # 3. Desenha a Tabela e Botão Excel dentro da pasta
            st.markdown("---")
            def to_excel(df_export):
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_export.to_excel(writer, index=False, sheet_name='Extrato')
                return output.getvalue()

            st.download_button(
                label=f"📥 Baixar Extrato: {fila}",
                data=to_excel(df_fila),
                file_name=f'extrato_{fila.replace(" ", "_").lower()}.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                key=f"dl_{fila}" # Chave única necessária quando há vários botões de download
            )
            
            st.dataframe(
                df_fila,
                column_config={
                    "Link Salesforce": st.column_config.LinkColumn("Acessar", display_text="Abrir"),
                    "Abertura": st.column_config.DateColumn("Abertura", format="DD/MM/YYYY"),
                    "ID do Caso": None
                },
                use_container_width=True,
                hide_index=True
            )
