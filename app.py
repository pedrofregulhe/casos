import streamlit as st
import pandas as pd
from simple_salesforce import Salesforce
import plotly.express as px

# --- 1. CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Dashboard de Casos - Planejamento", layout="wide")
st.title("📊 Visão de Casos em Aberto (Tipo: OA)")

# --- 2. CONEXÃO COM O SALESFORCE ---
# Usamos cache para não ficar logando no Salesforce a cada clique na tela
@st.cache_resource
def init_connection():
    return Salesforce(
        username=st.secrets["sf_username"],
        password=st.secrets["sf_password"],
        security_token=st.secrets["sf_token"],
        domain='login' # mude para 'test' se for sandbox
    )

sf = init_connection()

# --- 3. EXTRAÇÃO E TRATAMENTO DE DADOS ---
@st.cache_data(ttl=600) # Atualiza os dados a cada 10 minutos
def get_data():
    query = """
    SELECT 
        CaseNumber, Account.Name, Account.FOZ_StatusPosicaoFinanceira__c, Account.FOZ_CPF__c,
        Origin, Type, Status, FOZ_TipoSolicitacao__c, FOZ_Motivo__c, FOZ_Detalhe__c, FOZ_Subdetalhe__c, 
        Owner.Name, 
        (SELECT MilestoneType.Name, TargetDate, IsViolated FROM CaseMilestones)
    FROM Case 
    WHERE Type = 'OA' AND Status != 'Closed' AND Status != 'Fechado'
    """
    
    result = sf.query_all(query)
    
    # "Achatando" os dados (tirando do formato JSON aninhado)
    linhas = []
    for record in result['records']:
        # Tratando relacionamentos que podem vir vazios
        conta_nome = record['Account']['Name'] if record['Account'] else None
        dono_nome = record['Owner']['Name'] if record['Owner'] else 'Sem Dono'
        
        # Analisando o SLA (verificando se algum marco está violado)
        sla_violado = "No Prazo"
        if record['CaseMilestones'] and record['CaseMilestones']['records']:
            for milestone in record['CaseMilestones']['records']:
                if milestone['IsViolated']:
                    sla_violado = "Atrasado"
                    break # Se um estiver atrasado, o caso todo está atrasado
                    
        linhas.append({
            'Número do Caso': record['CaseNumber'],
            'Fila/Proprietário': dono_nome,
            'Conta': conta_nome,
            'Status': record['Status'],
            'Motivo': record['FOZ_Motivo__c'],
            'Detalhe': record['FOZ_Detalhe__c'],
            'Status SLA': sla_violado
        })
        
    return pd.DataFrame(linhas)

# Carrega o dataframe
df = get_data()

# --- 4. CONSTRUÇÃO DO DASHBOARD ---
if not df.empty:
    # Métricas (KPIs)
    col1, col2, col3 = st.columns(3)
    col1.metric("Total de Casos Abertos", len(df))
    
    casos_atrasados = len(df[df['Status SLA'] == 'Atrasado'])
    col2.metric("Casos com SLA Atrasado", casos_atrasados)
    
    filas_distintas = df['Fila/Proprietário'].nunique()
    col3.metric("Filas Atuantes", filas_distintas)
    
    st.markdown("---")
    
    # Gráficos
    col_grafico1, col_grafico2 = st.columns(2)
    
    with col_grafico1:
        st.subheader("Volume por Fila/Proprietário")
        fig_fila = px.bar(df['Fila/Proprietário'].value_counts().reset_index(), 
                          x='Fila/Proprietário', y='count', labels={'count': 'Volume de Casos'})
        st.plotly_chart(fig_fila, use_container_width=True)
        
    with col_grafico2:
        st.subheader("Saúde do SLA")
        fig_sla = px.pie(df, names='Status SLA', hole=0.4, color='Status SLA', 
                         color_discrete_map={'No Prazo':'#00CC96', 'Atrasado':'#EF553B'})
        st.plotly_chart(fig_sla, use_container_width=True)

    # Extrato de Dados
    st.markdown("---")
    st.subheader("Extrato Detalhado")
    # Filtro interativo para a tabela
    fila_selecionada = st.selectbox("Filtrar Tabela por Fila:", ["Todas"] + list(df['Fila/Proprietário'].unique()))
    if fila_selecionada != "Todas":
        st.dataframe(df[df['Fila/Proprietário'] == fila_selecionada], use_container_width=True)
    else:
        st.dataframe(df, use_container_width=True)
else:
    st.warning("Nenhum caso aberto encontrado no momento.")