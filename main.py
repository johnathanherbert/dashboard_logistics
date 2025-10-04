import streamlit as st
import pandas as pd
import io
import plotly.express as px

# --- Configuração da Página ---
st.set_page_config(
    page_title="Dashboard de Análise de Armazém",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- Funções Auxiliares ---
@st.cache_data
def load_data(uploaded_file):
    """
    Carrega e pré-processa os dados do arquivo Excel (.xls ou .xlsx) enviado.
    """
    if uploaded_file is None:
        return None
    
    try:
        # Tenta ler o arquivo como Excel (.xls ou .xlsx)
        df = pd.read_excel(uploaded_file)

        # --- Limpeza e Pré-processamento dos Dados ---
        
        # 1. Limpeza de Colunas: Remove colunas não nomeadas e espaços extras
        cols_to_drop = [col for col in df.columns if col.startswith('Unnamed:')]
        if cols_to_drop:
            df.drop(columns=cols_to_drop, inplace=True)
            
        df.columns = df.columns.str.strip() # Remove espaços em branco extras
        
        # 2. Validação de colunas essenciais
        required_cols = ['Altura', 'Estado Contentor']
        if not all(col in df.columns for col in required_cols):
            missing_cols = [col for col in required_cols if col not in df.columns]
            st.error(f"O arquivo enviado está faltando colunas necessárias: {missing_cols}")
            st.info("Verifique se as colunas estão nomeadas exatamente como: 'Altura' e 'Estado Contentor'")
            return None

        # 3. Limpeza 'Estado Contentor': Remove todas as aspas e espaços
        df['Estado Contentor'] = df['Estado Contentor'].astype(str).str.strip('"\' ').str.strip()

        # 4. Conversão 'Altura' para numérico
        df['Altura'] = (
            df['Altura'].astype(str)
            .str.strip('"\' ')
            .str.replace(',', '.', regex=False)
        )
        
        # Converte para float, 'coerce' (força) os erros para NaN
        df.dropna(subset=['Altura'], inplace=True)
        df['Altura'] = pd.to_numeric(df['Altura'], errors='coerce')

        
        # Limita as Alturas Apenas para os valores esperados (0.75m ou 1.50m)
        valid_heights = [0.75, 1.50]
        df = df[df['Altura'].isin(valid_heights)].copy()
        
        if df.empty:
            st.warning("O arquivo foi carregado, mas não contém dados válidos de 'Altura' (0.75 ou 1.50) após o pré-processamento.")
            return None

        return df
    
    except Exception as e:
        # Este bloco captura erros de I/O, formato e falta das bibliotecas openpyxl/xlrd
        st.error(f"Erro ao processar o arquivo. Verifique se as bibliotecas 'openpyxl' e 'xlrd' estão instaladas.")
        st.warning(f"Detalhe do erro: {e}")
        return None

def display_dashboard(df, total_posicoes_geral, total_posicoes_075, total_posicoes_150):
    """Gera e exibe o dashboard principal no Streamlit."""
    
    # --- FILTRAGEM E CÁLCULOS ---
    
    # Filtro: Manter apenas 'Armazenado' e 'Fora do Armazém' (conforme a regra de negócio)
    valid_status = ['Armazenado', 'Fora do Armazém']
    df_filtered = df[df['Estado Contentor'].isin(valid_status)].copy()
    
    if df_filtered.empty:
        st.warning("Nenhum dado encontrado com os status 'Armazenado' ou 'Fora do Armazém' após a filtragem.")
        return
        
    # DataFrames específicos para KPIs (agora só contêm dados relevantes)
    df_armazenado = df_filtered[df_filtered['Estado Contentor'] == 'Armazenado']
    df_fora = df_filtered[df_filtered['Estado Contentor'] == 'Fora do Armazém']

    # Resultados Agrupados
    results = {
        # Gerais
        'total_armazenado': len(df_armazenado),
        'total_fora_armazem': len(df_fora),
        # VAGAS VAZIAS (SALDO) = Posições Totais - Itens Armazenados (pode ser negativo)
        'vagas_vazias_geral': total_posicoes_geral - len(df_armazenado),
        
        # 0.75m
        'armazenado_075': len(df_armazenado[df_armazenado['Altura'] == 0.75]),
        'fora_armazem_075': len(df_fora[df_fora['Altura'] == 0.75]),
        # VAGAS VAZIAS 0.75m (SALDO)
        'vagas_vazias_075': total_posicoes_075 - len(df_armazenado[df_armazenado['Altura'] == 0.75]),
        
        # 1.50m
        'armazenado_150': len(df_armazenado[df_armazenado['Altura'] == 1.50]),
        'fora_armazem_150': len(df_fora[df_fora['Altura'] == 1.50]),
        # VAGAS VAZIAS 1.50m (SALDO)
        'vagas_vazias_150': total_posicoes_150 - len(df_armazenado[df_armazenado['Altura'] == 1.50]),
    }

    # Desempacota para facilitar o uso
    total_armazenado = results['total_armazenado']
    vagas_vazias_geral = results['vagas_vazias_geral']
    total_fora_armazem = results['total_fora_armazem']

    # Função auxiliar para formatar números com separador de milhar (ponto)
    def format_num(num):
        # Garante que números negativos sejam formatados corretamente (ex: -751)
        return f"{num:,}".replace(',', '.')

    # --- Definição de Estilos e Paleta Corporativa ---
    
    # Cores da Paleta Corporativa
    COR_VAZIO = '#0B72A4'       # Azul Corporativo (Disponibilidade)
    COR_OCUPADO = '#14854B'     # Verde Institucional (Ocupação)
    COR_FORA = '#B03A43'        # Vermelho Sóbrio (Alerta/Discrepância/Sobre-alocação)
    
    # Estilos CSS para as métricas
    common_style = "padding: 12px; border-radius: 8px; margin-bottom: 12px; border: 1px solid #ddd; height: 100%;" 
    style_vazio = f"border-left: 6px solid {COR_VAZIO}; background-color: rgba(11, 114, 164, 0.15); {common_style}"  
    style_ocupado = f"border-left: 6px solid {COR_OCUPADO}; background-color: rgba(20, 133, 75, 0.15); {common_style}" 
    style_fora = f"border-left: 6px solid {COR_FORA}; background-color: rgba(176, 58, 67, 0.15); {common_style}"    
    value_style = 'font-size: 1.8em; font-weight: 600;'


    # --- Exibição dos KPIs Principais ---
    st.header("Visão Geral do Armazém") 
    kpi_cols = st.columns(3)
    
    # KPI 1: Total Armazenado (Ocupação)
    kpi_cols[0].metric(
        "Ocupação Total", 
        format_num(total_armazenado), 
        help=f"Total de posições ocupadas fisicamente no armazém."
    )
    
    # KPI 2: Vagas Vazias (SALDO - Exibe Negativo se for o caso)
    
    kpi_title_vazias = "Vagas Vazias (Saldo)"
    kpi_value_vazias = format_num(vagas_vazias_geral) # Exibe o valor real, que será negativo se houver excesso
    
    if vagas_vazias_geral >= 0:
        disponibilidade_perc = (vagas_vazias_geral / total_posicoes_geral * 100) if total_posicoes_geral else 0
        delta_text_vazias = f"{disponibilidade_perc:.1f}% disponível"
        # Cor 'off' para valores neutros/positivos
        delta_color_vazias = "normal" if disponibilidade_perc >= 50 else "off" 
    else:
        # Se negativo, indica sobre-alocação no delta
        delta_text_vazias = f"Sobre-alocação: {format_num(abs(vagas_vazias_geral))} posições"
        delta_color_vazias = "inverse" # Vermelho para sobre-alocação

    kpi_cols[1].metric(
        kpi_title_vazias, 
        kpi_value_vazias, 
        delta=delta_text_vazias, 
        delta_color=delta_color_vazias 
    )
    
    # KPI 3: Fora do Armazém (Discrepância / Atenção)
    kpi_cols[2].metric(
        "Discrepância (Fora do Armazém)", 
        format_num(total_fora_armazem), 
        help="Itens registrados como 'Fora do Armazém', excluindo outros status como 'Em Trânsito' ou 'Lost and Found'.", 
        delta_color="inverse"
    )
    
    # --- LEGENDA SOLICITADA PARA O KPI DE DISCREPÂNCIA ---
    st.caption(f"""
        <div style="padding: 10px; border-radius: 5px; background-color: #f7f7f7;">
            <span style="color: {COR_FORA}; font-weight: bold;">ⓘ Nota sobre Discrepância:</span> O valor 'Fora do Armazém' inclui apenas itens com status 
            'Fora do Armazém' no arquivo de dados. Outros status temporários são ignorados para focar em problemas de inventário.
        </div>
    """, unsafe_allow_html=True)

    st.divider()

    # --- Análise por Altura (Metrics) ---
    st.header("Análise Detalhada por Posição")
    st.caption(f"Posições de 0.75m: {format_num(total_posicoes_075)} | Posições de 1.50m: {format_num(total_posicoes_150)}")

    metric_detail_cols = st.columns(6, gap="small")

    # 0.75m
    vazias_075 = results["vagas_vazias_075"]
    vazias_075_display = format_num(vazias_075) # Exibe o valor real (pode ser negativo)
    # Título dinâmico: "Vazias" se positivo/zero, "Sobre-alocação" se negativo
    vazias_075_title = "0.75m Vazias (Saldo)" if vazias_075 >= 0 else "0.75m Sobre-alocação"
    # Determina o estilo: Azul (Vazio) se >= 0, Vermelho (Alerta) se < 0
    style_075 = style_vazio if vazias_075 >= 0 else style_fora 
    
    metric_detail_cols[0].markdown(f'<div style="{style_075}">{vazias_075_title}<br><span style="{value_style}">{vazias_075_display}</span></div>', unsafe_allow_html=True)
    metric_detail_cols[1].markdown(f'<div style="{style_ocupado}">0.75m Armazenado<br><span style="{value_style}">{format_num(results["armazenado_075"])}</span></div>', unsafe_allow_html=True)
    metric_detail_cols[2].markdown(f'<div style="{style_fora}">0.75m Fora do Armazém<br><span style="{value_style}">{format_num(results["fora_armazem_075"])}</span></div>', unsafe_allow_html=True)
    
    # 1.50m
    vazias_150 = results["vagas_vazias_150"]
    vazias_150_display = format_num(vazias_150) # Exibe o valor real (pode ser negativo)
    # Título dinâmico: "Vazias" se positivo/zero, "Sobre-alocação" se negativo
    vazias_150_title = "1.50m Vazias (Saldo)" if vazias_150 >= 0 else "1.50m Sobre-alocação"
    # Determina o estilo: Azul (Vazio) se >= 0, Vermelho (Alerta) se < 0
    style_150 = style_vazio if vazias_150 >= 0 else style_fora
    
    metric_detail_cols[3].markdown(f'<div style="{style_150}">{vazias_150_title}<br><span style="{value_style}">{vazias_150_display}</span></div>', unsafe_allow_html=True)
    metric_detail_cols[4].markdown(f'<div style="{style_ocupado}">1.50m Armazenado<br><span style="{value_style}">{format_num(results["armazenado_150"])}</span></div>', unsafe_allow_html=True)
    metric_detail_cols[5].markdown(f'<div style="{style_fora}">1.50m Fora do Armazém<br><span style="{value_style}">{format_num(results["fora_armazem_150"])}</span></div>', unsafe_allow_html=True)


    # --- Gráficos de Pizza Detalhados (PIE CHARTS) ---
    st.divider()
    st.header("Distribuição de Ocupação por Altura")
    st.caption("Proporção entre posições Ocupadas (Armazenado) e Vagas (Vazio/Excesso de Ocupação).")

    # Data preparation for the pie charts
    
    # Chart 1: 0.75m (Usa 'Excesso de Ocupação' para valores negativos)
    pie_status_075 = 'Vazio' if vazias_075 >= 0 else 'Excesso de Ocupação'
    df_pie_075 = pd.DataFrame({
        'Status': ['Armazenado', pie_status_075],
        # Usamos o valor absoluto do VAZIO/EXCESSO para o gráfico
        'Quantidade': [results['armazenado_075'], abs(vazias_075)]
    })
    
    # Chart 2: 1.50m
    pie_status_150 = 'Vazio' if vazias_150 >= 0 else 'Excesso de Ocupação'
    df_pie_150 = pd.DataFrame({
        'Status': ['Armazenado', pie_status_150],
        'Quantidade': [results['armazenado_150'], abs(vazias_150)]
    })

    chart_col1, chart_col2 = st.columns(2)

    with chart_col1:
        # Gráfico de Pizza 1: Ocupação 0.75m
        fig_pie_075 = px.pie(
            df_pie_075,
            names='Status',
            values='Quantidade',
            title=f'Ocupação de Posições 0.75m (Ref. Capacidade: {format_num(total_posicoes_075)})',
            color='Status',
            color_discrete_map={
                'Armazenado': COR_OCUPADO,
                'Vazio': COR_VAZIO,
                'Excesso de Ocupação': COR_FORA # Usa a cor de alerta para Excesso de Ocupação
            },
            template='plotly_white',
            hole=0.4, # Transforma em Donut Chart
        )
        fig_pie_075.update_traces(
            textinfo='percent+value', 
            textposition='inside',
            marker=dict(line=dict(color='#FFFFFF', width=2)),
            hovertemplate='<b>%{label}</b><br>Quantidade: %{value}<br>Percentual: %{percent}<extra></extra>'
        )
        fig_pie_075.update_layout(
            title=dict(x=0.5),
            legend_title="Status da Posição",
            margin=dict(t=50, b=20, l=20, r=20),
            paper_bgcolor='rgba(0,0,0,0)', 
            plot_bgcolor='rgba(0,0,0,0)',
        )
        st.plotly_chart(fig_pie_075, use_container_width=True)

    with chart_col2:
        # Gráfico de Pizza 2: Ocupação 1.50m
        fig_pie_150 = px.pie(
            df_pie_150,
            names='Status',
            values='Quantidade',
            title=f'Ocupação de Posições 1.50m (Ref. Capacidade: {format_num(total_posicoes_150)})',
            color='Status',
            color_discrete_map={
                'Armazenado': COR_OCUPADO,
                'Vazio': COR_VAZIO,
                'Excesso de Ocupação': COR_FORA # Usa a cor de alerta para Excesso de Ocupação
            },
            template='plotly_white',
            hole=0.4, # Transforma em Donut Chart
        )
        fig_pie_150.update_traces(
            textinfo='percent+value',
            textposition='inside',
            marker=dict(line=dict(color='#FFFFFF', width=2)),
            hovertemplate='<b>%{label}</b><br>Quantidade: %{value}<br>Percentual: %{percent}<extra></extra>'
        )
        fig_pie_150.update_layout(
            title=dict(x=0.5),
            legend_title="Status da Posição",
            margin=dict(t=50, b=20, l=20, r=20),
            paper_bgcolor='rgba(0,0,0,0)', 
            plot_bgcolor='rgba(0,0,0,0)',
        )
        st.plotly_chart(fig_pie_150, use_container_width=True)


    # --- Gráfico de Alerta de Discrepância (Potential Put-Away) ---
    st.subheader("Itens Registrados como 'Fora do Armazém' (Potencial a Armazenar)")

    discrepancy_data = {
        'Altura': ['0.75m', '1.50m'],
        'Fora do Armazém': [results['fora_armazem_075'], results['fora_armazem_150']]
    }
    df_discrepancy = pd.DataFrame(discrepancy_data)

    fig_discrepancy = px.bar(
        df_discrepancy,
        x='Altura',
        y='Fora do Armazém',
        title='Contagem de Itens Fora do Armazém por Altura (Potencial de Entrada)', # Título ajustado para clareza
        color='Altura',
        color_discrete_sequence=[COR_FORA, '#D36A72'],
        text_auto=True,
        template='plotly_white',
        hover_data={'Fora do Armazém': True, 'Altura': False}
    )
    fig_discrepancy.update_layout(
        xaxis_title="Altura da Posição",
        yaxis_title="Número de Posições",
        showlegend=False,
        margin=dict(t=50, b=20, l=20, r=20),
        paper_bgcolor='rgba(0,0,0,0)', 
        plot_bgcolor='rgba(0,0,0,0)',
        font=dict(size=12),
        title=dict(x=0.5), 
        hovermode="x unified"
    )
    fig_discrepancy.update_xaxes(showline=True, linewidth=1, linecolor='lightgrey')
    fig_discrepancy.update_yaxes(showgrid=True, gridwidth=1, gridcolor='lightgrey')
    st.plotly_chart(fig_discrepancy, use_container_width=True)


    st.divider()

    # --- Tabela de Dados Interativa ---
    st.header("Explore os Dados Detalhados (Apenas Armazenado e Fora do Armazém)")
    st.dataframe(df_filtered, use_container_width=True, height=400, key='data_explorer')


# --- Lógica Principal (Execução) ---

# Interface do Streamlit (Sidebar)
with st.sidebar:
    st.image("https://gruponc.net.br/wp-content/webp-express/webp-images/uploads/2024/06/logo-novamed.png.webp", width=100)
    st.title("Configurações do Armazém")
    st.markdown("Ajuste os parâmetros do seu armazém e envie seu arquivo de dados.")

    total_posicoes_geral = st.number_input(
        "Total de Posições no AVR",
        min_value=1,
        value=4060,
        step=10,
        help="Informe o número total de locais de armazenamento disponíveis no armazém."
    )

    st.subheader("Distribuição de Posições por Altura")
    total_posicoes_075 = st.number_input("Total de Posições de 0.75m", min_value=0, value=2030, step=10)
    total_posicoes_150 = st.number_input("Total de Posições de 1.50m", min_value=0, value=2030, step=10)

    if total_posicoes_075 + total_posicoes_150 != total_posicoes_geral:
        st.warning("A soma das posições de 0.75m e 1.50m não corresponde ao total geral de posições.")

    uploaded_file = st.file_uploader(
        "Carregue seu arquivo de dados (Excel)",
        type=["xls", "xlsx"] 
    )
    st.info("O arquivo deve ser um XLS ou XLSX (Excel) com as colunas 'Altura' e 'Estado Contentor' na primeira aba.")


# Título Principal
st.title("Dashboard de Análise de Armazém")
st.markdown("Visão geral da ocupação, vagas disponíveis e discrepâncias no armazenamento.")


if uploaded_file is None:
    st.info("Aguardando o envio do arquivo de dados para iniciar a análise.")
else:
    # Chama a função de carregamento (usa @st.cache_data para performance)
    df = load_data(uploaded_file)

    if df is not None:
        # Chama a função de exibição do dashboard
        display_dashboard(df, total_posicoes_geral, total_posicoes_075, total_posicoes_150)
