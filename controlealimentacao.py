import streamlit as st
import requests
from msal import ConfidentialClientApplication
import pandas as pd
import io
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from datetime import datetime, timedelta
import numpy as np
import openpyxl

# Configuração da página
st.set_page_config(
    page_title="Painel Gerencial - Alimentações",
    page_icon="🍽️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# CSS customizado com cores da empresa
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(90deg, #F7931E 0%, #000000 100%);
        padding: 2rem;
        border-radius: 10px;
        margin-bottom: 2rem;
        text-align: center;
        color: white;
    }

    .metric-card {
        background: white;
        padding: 1.5rem;
        border-radius: 10px;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        border-left: 4px solid #F7931E;
        margin-bottom: 1rem;
    }

    .metric-title {
        font-size: 0.9rem;
        color: #666;
        margin-bottom: 0.5rem;
    }

    .metric-value {
        font-size: 2rem;
        font-weight: bold;
        color: #000000;
        margin: 0;
    }

    .metric-delta {
        font-size: 0.8rem;
        margin-top: 0.5rem;
    }

    .sidebar .sidebar-content {
        background: linear-gradient(180deg, #F7931E 0%, #000000 100%);
    }

    .stSelectbox > div > div > div {
        background-color: #f8f9fa;
    }

    .chart-container {
        background: white;
        padding: 1rem;
        border-radius: 10px;
        box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
        margin-bottom: 1rem;
    }

    .success-message {
        background: #d4edda;
        border: 1px solid #c3e6cb;
        color: #155724;
        padding: 1rem;
        border-radius: 5px;
        margin: 1rem 0;
    }

    .error-message {
        background: #f8d7da;
        border: 1px solid #f5c6cb;
        color: #721c24;
        padding: 1rem;
        border-radius: 5px;
        margin: 1rem 0;
    }

    /* Estilização adicional com cores da empresa */
    .stTab [data-baseweb="tab-list"] {
        gap: 8px;
    }

    .stTab [data-baseweb="tab"] {
        background-color: #F7931E;
        color: white;
        border-radius: 8px;
    }

    .stTab [aria-selected="true"] {
        background-color: #000000 !important;
        color: white;
    }
</style>
""", unsafe_allow_html=True)


@st.cache_data(ttl=300)  # Cache por 5 minutos
def download_excel_sharepoint():
    """Baixa dados do SharePoint usando st.secrets"""
    try:
        # Configurar autenticação usando st.secrets
        app = ConfidentialClientApplication(
            st.secrets["azure"]["client_id"],
            authority=f"https://login.microsoftonline.com/{st.secrets['azure']['tenant_id']}",
            client_credential=st.secrets["azure"]["client_secret"],
        )

        # Obter token
        result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])

        if "access_token" in result:
            headers = {"Authorization": f"Bearer {result['access_token']}"}

            # Obter o site_id
            site_url = "https://graph.microsoft.com/v1.0/sites/rezendeenergia.sharepoint.com:/sites/Intranet"
            site_response = requests.get(site_url, headers=headers)

            if site_response.status_code == 200:
                site_data = site_response.json()
                site_id = site_data['id']

                # Buscar o arquivo específico
                search_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root/search(q='Controle Alimentação.xlsx')"
                search_response = requests.get(search_url, headers=headers)

                if search_response.status_code == 200:
                    search_data = search_response.json()
                    files_found = search_data.get('value', [])

                    for item in files_found:
                        if item['name'] == 'Controle Alimentação.xlsx':
                            # Baixar o arquivo
                            download_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/items/{item['id']}/content"
                            download_response = requests.get(download_url, headers=headers)

                            if download_response.status_code == 200:
                                # Ler o arquivo Excel
                                df = pd.read_excel(io.BytesIO(download_response.content))
                                return df

        return None

    except Exception as e:
        st.error(f"Erro ao conectar com SharePoint: {e}")
        return None


def process_data(df):
    """Processa e limpa os dados"""
    if df is None:
        return None

    # Renomear colunas para facilitar o trabalho
    df.columns = ['data_compra', 'item', 'unidade_medida', 'valor_unitario',
                  'quantidade', 'valor_total', 'categoria', 'alojamento']

    # Converter tipos de dados
    df['data_compra'] = pd.to_datetime(df['data_compra'])
    df['valor_unitario'] = pd.to_numeric(df['valor_unitario'], errors='coerce')
    df['quantidade'] = pd.to_numeric(df['quantidade'], errors='coerce')
    df['valor_total'] = pd.to_numeric(df['valor_total'], errors='coerce')

    # Adicionar colunas calculadas
    df['mes_ano'] = df['data_compra'].dt.to_period('M')
    df['dia_semana'] = df['data_compra'].dt.day_name()
    df['semana'] = df['data_compra'].dt.isocalendar().week

    return df


def create_metrics_cards(df, col1, col2, col3, col4):
    """Cria cards de métricas principais"""

    # Métricas gerais
    total_gasto = df['valor_total'].sum()
    total_itens = len(df)
    gasto_medio_dia = df.groupby('data_compra')['valor_total'].sum().mean()
    alojamentos_ativos = df['alojamento'].nunique()

    # Métricas do mês atual vs anterior
    hoje = datetime.now()
    mes_atual = df[df['data_compra'].dt.month == hoje.month]
    mes_anterior = df[df['data_compra'].dt.month == (hoje.month - 1 if hoje.month > 1 else 12)]

    gasto_mes_atual = mes_atual['valor_total'].sum() if len(mes_atual) > 0 else 0
    gasto_mes_anterior = mes_anterior['valor_total'].sum() if len(mes_anterior) > 0 else 0

    if gasto_mes_anterior > 0:
        variacao_mensal = ((gasto_mes_atual - gasto_mes_anterior) / gasto_mes_anterior) * 100
    else:
        variacao_mensal = 0

    with col1:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-title">💰 Gasto Total</div>
            <div class="metric-value">R$ {total_gasto:,.2f}</div>
        </div>
        """, unsafe_allow_html=True)

    with col2:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-title">📦 Total de Itens</div>
            <div class="metric-value">{total_itens:,}</div>
        </div>
        """, unsafe_allow_html=True)

    with col3:
        variacao_color = "#F7931E" if variacao_mensal >= 0 else "#FF0000"
        variacao_icon = "↗️" if variacao_mensal >= 0 else "↘️"
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-title">📅 Gasto Médio/Dia</div>
            <div class="metric-value">R$ {gasto_medio_dia:.2f}</div>
            <div class="metric-delta" style="color: {variacao_color}">
                {variacao_icon} {variacao_mensal:.1f}% vs mês anterior
            </div>
        </div>
        """, unsafe_allow_html=True)

    with col4:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-title">🏠 Alojamentos Ativos</div>
            <div class="metric-value">{alojamentos_ativos}</div>
        </div>
        """, unsafe_allow_html=True)


def create_charts(df):
    """Cria gráficos do dashboard"""

    # Cores da empresa
    cores_empresa = ['#F7931E', '#000000', '#FF6B35', '#FFB366', '#333333', '#666666', '#999999']

    col1, col2 = st.columns(2)

    with col1:
        # Gráfico de gastos por categoria
        gastos_categoria = df.groupby('categoria')['valor_total'].sum().reset_index()
        gastos_categoria = gastos_categoria.sort_values('valor_total', ascending=False)

        fig_categoria = px.pie(
            gastos_categoria,
            values='valor_total',
            names='categoria',
            title="💳 Distribuição de Gastos por Categoria",
            color_discrete_sequence=cores_empresa
        )
        fig_categoria.update_layout(
            height=400,
            title_x=0.5,
            font=dict(size=12),
            title_font_color='#000000'
        )
        st.plotly_chart(fig_categoria, use_container_width=True)

    with col2:
        # Gráfico de gastos por alojamento
        gastos_alojamento = df.groupby('alojamento')['valor_total'].sum().reset_index()
        gastos_alojamento = gastos_alojamento.sort_values('valor_total', ascending=True)

        fig_alojamento = px.bar(
            gastos_alojamento,
            x='valor_total',
            y='alojamento',
            title="🏠 Gastos por Alojamento",
            orientation='h',
            color='valor_total',
            color_continuous_scale=[[0, '#F7931E'], [1, '#000000']]
        )
        fig_alojamento.update_layout(
            height=400,
            title_x=0.5,
            showlegend=False,
            title_font_color='#000000'
        )
        st.plotly_chart(fig_alojamento, use_container_width=True)

    # Gráfico de evolução temporal
    gastos_diarios = df.groupby('data_compra')['valor_total'].sum().reset_index()

    fig_timeline = px.line(
        gastos_diarios,
        x='data_compra',
        y='valor_total',
        title="📈 Evolução dos Gastos ao Longo do Tempo",
        line_shape='spline'
    )
    fig_timeline.update_traces(line_color='#F7931E', line_width=3)
    fig_timeline.update_layout(
        height=400,
        title_x=0.5,
        xaxis_title="Data",
        yaxis_title="Valor Total (R$)",
        title_font_color='#000000'
    )
    st.plotly_chart(fig_timeline, use_container_width=True)

    # Heatmap de gastos por dia da semana e categoria
    df_pivot = df.pivot_table(
        values='valor_total',
        index='categoria',
        columns='dia_semana',
        aggfunc='sum',
        fill_value=0
    )

    # Reordenar dias da semana
    dias_ordem = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
    df_pivot = df_pivot.reindex(columns=[dia for dia in dias_ordem if dia in df_pivot.columns])

    fig_heatmap = px.imshow(
        df_pivot.values,
        x=[dia[:3] for dia in df_pivot.columns],
        y=df_pivot.index,
        title="🗓️ Heatmap: Gastos por Categoria e Dia da Semana",
        color_continuous_scale=[[0, '#FFFFFF'], [0.5, '#F7931E'], [1, '#000000']],
        aspect='auto'
    )
    fig_heatmap.update_layout(
        height=400,
        title_x=0.5,
        xaxis_title="Dia da Semana",
        yaxis_title="Categoria",
        title_font_color='#000000'
    )
    st.plotly_chart(fig_heatmap, use_container_width=True)


def create_detailed_analysis(df):
    """Cria análises detalhadas"""

    st.markdown("## 🔍 Análise Detalhada")

    tab1, tab2, tab3, tab4 = st.tabs(["📊 Top Produtos", "💰 Análise Financeira", "🏠 Por Alojamento", "📅 Tendências"])

    with tab1:
        col1, col2 = st.columns(2)

        with col1:
            # Top produtos mais comprados
            top_produtos = df.groupby('item').agg({
                'quantidade': 'sum',
                'valor_total': 'sum'
            }).sort_values('quantidade', ascending=False).head(10)

            st.markdown("### 📈 Top 10 - Produtos Mais Comprados")
            for idx, (produto, row) in enumerate(top_produtos.iterrows(), 1):
                st.markdown(f"""
                <div style="padding: 0.5rem; border-left: 3px solid #F7931E; margin-bottom: 0.5rem; background: #f8f9fa;">
                    <strong>{idx}. {produto}</strong><br>
                    Quantidade: {row['quantidade']:.0f} | Valor: R$ {row['valor_total']:,.2f}
                </div>
                """, unsafe_allow_html=True)

        with col2:
            # Produtos mais caros
            produtos_caros = df.groupby('item')['valor_unitario'].mean().sort_values(ascending=False).head(10)

            st.markdown("### 💎 Top 10 - Produtos Mais Caros (Valor Unitário)")
            for idx, (produto, valor) in enumerate(produtos_caros.items(), 1):
                st.markdown(f"""
                <div style="padding: 0.5rem; border-left: 3px solid #000000; margin-bottom: 0.5rem; background: #f8f9fa;">
                    <strong>{idx}. {produto}</strong><br>
                    Valor Unitário: R$ {valor:,.2f}
                </div>
                """, unsafe_allow_html=True)

    with tab2:
        col1, col2 = st.columns(2)

        with col1:
            # Análise por mês
            gastos_mes = df.groupby('mes_ano')['valor_total'].sum().reset_index()
            gastos_mes['mes_ano_str'] = gastos_mes['mes_ano'].astype(str)

            fig_mes = px.bar(
                gastos_mes,
                x='mes_ano_str',
                y='valor_total',
                title="📅 Gastos por Mês",
                color='valor_total',
                color_continuous_scale=[[0, '#F7931E'], [1, '#000000']]
            )
            fig_mes.update_layout(title_font_color='#000000')
            st.plotly_chart(fig_mes, use_container_width=True)

        with col2:
            # Distribuição de valores
            fig_dist = px.histogram(
                df,
                x='valor_total',
                nbins=30,
                title="📊 Distribuição de Valores das Compras",
                color_discrete_sequence=['#F7931E']
            )
            fig_dist.update_layout(title_font_color='#000000')
            st.plotly_chart(fig_dist, use_container_width=True)

    with tab3:
        # Análise por alojamento
        alojamento_stats = df.groupby('alojamento').agg({
            'valor_total': ['sum', 'mean', 'count'],
            'quantidade': 'sum'
        }).round(2)

        alojamento_stats.columns = ['Total Gasto', 'Gasto Médio', 'Nº Compras', 'Quantidade Total']
        alojamento_stats = alojamento_stats.sort_values('Total Gasto', ascending=False)

        st.markdown("### 🏠 Estatísticas por Alojamento")
        st.dataframe(alojamento_stats, use_container_width=True)

        # Gráfico de comparação
        fig_aloj_comp = go.Figure()

        fig_aloj_comp.add_trace(go.Bar(
            name='Total Gasto',
            x=alojamento_stats.index,
            y=alojamento_stats['Total Gasto'],
            yaxis='y',
            offsetgroup=1,
            marker_color='#F7931E'
        ))

        fig_aloj_comp.add_trace(go.Bar(
            name='Nº Compras',
            x=alojamento_stats.index,
            y=alojamento_stats['Nº Compras'],
            yaxis='y2',
            offsetgroup=2,
            marker_color='#000000'
        ))

        fig_aloj_comp.update_layout(
            title="📊 Comparativo: Gastos vs Número de Compras por Alojamento",
            xaxis_title="Alojamento",
            yaxis=dict(title="Valor Total (R$)", side="left"),
            yaxis2=dict(title="Número de Compras", side="right", overlaying="y"),
            height=500,
            title_font_color='#000000'
        )

        st.plotly_chart(fig_aloj_comp, use_container_width=True)

    with tab4:
        col1, col2 = st.columns(2)

        with col1:
            # Gastos por dia da semana - CORREÇÃO DO BUG
            gastos_dia_semana = df.groupby('dia_semana')['valor_total'].mean().reset_index()
            ordem_dias = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']

            # Filtrar apenas os dias que existem nos dados
            dias_presentes = gastos_dia_semana['dia_semana'].tolist()
            gastos_dia_semana['dia_num'] = gastos_dia_semana['dia_semana'].apply(lambda x: ordem_dias.index(x))
            gastos_dia_semana = gastos_dia_semana.sort_values('dia_num')

            # Mapear apenas os dias presentes
            mapeamento_dias = {
                'Monday': 'Seg', 'Tuesday': 'Ter', 'Wednesday': 'Qua',
                'Thursday': 'Qui', 'Friday': 'Sex', 'Saturday': 'Sáb', 'Sunday': 'Dom'
            }
            gastos_dia_semana['dia_pt'] = gastos_dia_semana['dia_semana'].map(mapeamento_dias)

            fig_dia_semana = px.bar(
                gastos_dia_semana,
                x='dia_pt',
                y='valor_total',
                title="📅 Gasto Médio por Dia da Semana",
                color='valor_total',
                color_continuous_scale=[[0, '#F7931E'], [1, '#000000']],
                text='valor_total'
            )
            fig_dia_semana.update_traces(texttemplate='R$ %{text:,.0f}', textposition='outside')
            fig_dia_semana.update_layout(title_font_color='#000000')
            st.plotly_chart(fig_dia_semana, use_container_width=True)

        with col2:
            # Sazonalidade mensal
            if len(df) > 0:
                df['mes'] = df['data_compra'].dt.month
                gastos_sazonalidade = df.groupby('mes')['valor_total'].mean().reset_index()
                meses_nomes = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun',
                               'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez']
                gastos_sazonalidade['mes_nome'] = gastos_sazonalidade['mes'].apply(lambda x: meses_nomes[x - 1])

                fig_sazonalidade = px.line(
                    gastos_sazonalidade,
                    x='mes_nome',
                    y='valor_total',
                    title="🌟 Sazonalidade - Gasto Médio por Mês",
                    markers=True
                )
                fig_sazonalidade.update_traces(line_color='#F7931E', line_width=3, marker_size=8,
                                               marker_color='#000000')
                fig_sazonalidade.update_layout(title_font_color='#000000')
                st.plotly_chart(fig_sazonalidade, use_container_width=True)
            else:
                st.info("📊 Dados insuficientes para análise de sazonalidade")


def main():
    """Função principal do dashboard"""

    # Header principal
    st.markdown("""
    <div class="main-header">
        <h1>🍽️ Painel Gerencial - Controle de Alimentações</h1>
        <p>Análise completa dos gastos e consumo por alojamento</p>
    </div>
    """, unsafe_allow_html=True)

    # Sidebar para filtros
    with st.sidebar:
        st.markdown("## 🔧 Filtros e Configurações")

        # Botão para atualizar dados
        if st.button("🔄 Atualizar Dados", use_container_width=True):
            st.cache_data.clear()
            st.rerun()

        st.markdown("---")

        # Carregar dados
        with st.spinner("📊 Carregando dados do SharePoint..."):
            df = download_excel_sharepoint()

        if df is not None:
            df = process_data(df)
            st.success(f"✅ {len(df)} registros carregados!")

            # Filtros
            st.markdown("### 📅 Período")
            data_min = df['data_compra'].min().date()
            data_max = df['data_compra'].max().date()

            data_inicio, data_fim = st.date_input(
                "Selecione o período:",
                value=[data_min, data_max],
                min_value=data_min,
                max_value=data_max
            )

            # Filtro por alojamento
            st.markdown("### 🏠 Alojamentos")
            alojamentos_disponiveis = ['Todos'] + sorted(df['alojamento'].unique().tolist())
            alojamento_selecionado = st.selectbox(
                "Selecione o alojamento:",
                alojamentos_disponiveis
            )

            # Filtro por categoria
            st.markdown("### 📦 Categorias")
            categorias_disponiveis = ['Todas'] + sorted(df['categoria'].unique().tolist())
            categoria_selecionada = st.selectbox(
                "Selecione a categoria:",
                categorias_disponiveis
            )

            # Aplicar filtros
            df_filtrado = df[
                (df['data_compra'].dt.date >= data_inicio) &
                (df['data_compra'].dt.date <= data_fim)
                ]

            if alojamento_selecionado != 'Todos':
                df_filtrado = df_filtrado[df_filtrado['alojamento'] == alojamento_selecionado]

            if categoria_selecionada != 'Todas':
                df_filtrado = df_filtrado[df_filtrado['categoria'] == categoria_selecionada]

            st.markdown(f"**📊 {len(df_filtrado)} registros após filtros**")

        else:
            st.error("❌ Não foi possível carregar os dados do SharePoint.")
            st.stop()

    # Dashboard principal
    if df_filtrado is not None and len(df_filtrado) > 0:

        # Métricas principais
        col1, col2, col3, col4 = st.columns(4)
        create_metrics_cards(df_filtrado, col1, col2, col3, col4)

        # Gráficos principais
        st.markdown("## 📈 Visualizações")
        create_charts(df_filtrado)

        # Análise detalhada
        create_detailed_analysis(df_filtrado)

        # Tabela de dados brutos
        with st.expander("📋 Dados Detalhados"):
            st.dataframe(
                df_filtrado.sort_values('data_compra', ascending=False),
                use_container_width=True,
                hide_index=True
            )

            # Opção para download
            csv = df_filtrado.to_csv(index=False)
            st.download_button(
                label="💾 Baixar dados filtrados (CSV)",
                data=csv,
                file_name=f"alimentacoes_filtrado_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                mime="text/csv"
            )

        # Rodapé
        st.markdown("---")
        st.markdown(
            """
            <div style="text-align: center; color: #666; padding: 1rem;">
                🍽️ Painel Gerencial de Alimentações | Desenvolvido com Streamlit<br>
                Última atualização: """ + datetime.now().strftime("%d/%m/%Y às %H:%M") + """
            </div>
            """,
            unsafe_allow_html=True
        )

    else:
        st.warning("⚠️ Nenhum dado disponível com os filtros selecionados.")


if __name__ == "__main__":

    main()
