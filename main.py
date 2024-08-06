import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st
import os
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from io import BytesIO
from streamlit_gsheets import GSheetsConnection
from dotenv import load_dotenv

# Setup
## Carregar environment variables de .env file
load_dotenv()

## Acessar chave do API do google sheets
GOOGLE_SHEETS_API_KEY = os.getenv('GOOGLE_SHEETS_API_KEY')

## Configurar web page
st.set_page_config(
    page_title = 'Premier Pack - Dashboard',
    page_icon = ':bar_chart:',
    layout = 'wide'
)

## Conectar ao Google Sheets
ttl = 60 # time (seconds) it takes for data to be cleared from cache
conn = st.experimental_connection('gsheets', type=GSheetsConnection, api_key=GOOGLE_SHEETS_API_KEY)
df = conn.read(ttl=10)
#st.write(df)

# Funções
def df_to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Banco de Dados')
    output.seek(0)
    
    # Load workbook and worksheet
    wb = load_workbook(output)
    ws = wb['Banco de Dados']
    
    # Apply styles to header
    header_font = Font(bold=True, color='FFFFFF')
    header_fill = PatternFill(start_color='4F81BD', end_color='4F81BD', fill_type='solid')
    
    for cell in ws[1]:
        cell.font = header_font
        cell.fill = header_fill
        
    # Save workbook back
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    
    return output

## Converter data para datetime
df['DATA'] = pd.to_datetime(df['DATA'], format='%Y-%m-%d')

## Título
st.title(':bar_chart: Dashboard Premier Pack')

## Tabs
tabs = st.tabs(['Estatísticas', 'Maiores Faturamentos', 'Tabela', 'Baixar Folha'])

with tabs[0]:
    # Sidebar
    ## Definir filtros
    st.sidebar.header('Filtros:')

    if 'show_more_filters' not in st.session_state:
        st.session_state.show_more_filters = False
        
    def toggle_filters():
        st.session_state.show_more_filters = not st.session_state.show_more_filters

    ### Faturamento
    faturamento_min = st.sidebar.number_input('Fat. Mín.:', value=0)
    faturamento_max = st.sidebar.number_input('Fat. Máx.:', value=df['FATURAMENTO'].max())

    if faturamento_min > faturamento_max:
        st.sidebar.error('O faturamento mínimo tem de ser menor que o faturamento máximo.')

    ### Data
    start_date = pd.to_datetime(st.sidebar.date_input('Data Inicial', value=df['DATA'].min()))
    end_date   = pd.to_datetime(st.sidebar.date_input('Data Final',   value=df['DATA'].max()))

    if start_date > end_date:
        st.sidebar.error('A data final tem que ser anterior à data inicial.')

    ### Categóricos
    filters = {
        'CLIENTE' : st.sidebar.multiselect(
         'Clientes:',
         options = df['CLIENTE'].unique(),
         default = []
         ),
        
        'COMERCIAL' : st.sidebar.multiselect(
         'Comercial:',
         options = df['COMERCIAL'].unique(),
         default = []
        ),
        
        'TIPO DE PRODUTO' : st.sidebar.multiselect(
         'Tipo de Produto:',
         options = df['TIPO DE PRODUTO'].unique(),
         default = []
        ),
        
        'SEGMENTO' : st.sidebar.multiselect(
         'Segmento:',
         options = df['SEGMENTO'].unique(),
         default = []
        ),
        
        'MERCADO' : st.sidebar.multiselect(
         'Mercado:',
         options = df['MERCADO'].unique(),
         default = []
        ),
        
        'ESTADO' : st.sidebar.multiselect(
         'Estado:',
         options = df['UF'].unique(),
         default = []
        ),
        
        'PAÍS' : st.sidebar.multiselect(
         'País:',
         options = df['PAÍS'].unique(),
         default = []
        )
    }

    ## Filtros Opcionais
    st.sidebar.button('Mais Filtros...', on_click=toggle_filters)
    if st.session_state.show_more_filters:
        st.sidebar.subheader('Filtros Adicionais')
        opt_filters = {
            'MUNICÍPIO' : st.sidebar.multiselect(
             'Município:',
             options = df['MUNICÍPIO'].unique(),
             default = []
            ),
            
            'CONTINENTE' : st.sidebar.multiselect(
             'Continente:',
             options = df['CONTINENTE'].unique(),
             default = []
            ),
            
            'ICMS' : st.sidebar.multiselect(
             'Dentro/Fora do Estado:',
             options = df['ICMS'].unique(),
             default = []
            ),
            
            'PRODUTO' : st.sidebar.multiselect(
             'Produto:',
             options = df['PRODUTO'].unique(),
             default = []
            ),
            
            'ORIGEM DO PRODUTO' : st.sidebar.multiselect(
             'Origem do Produto:',
             options = df['ORIGEM DO PRODUTO'].unique(),
             default = []
            ),
        }  

    ## Aplicar filtros
    df_filtered = df

    for column, selected_values in filters.items():
        if selected_values:
            df_filtered = df_filtered[df_filtered[column].isin(selected_values)]

    df_filtered = df_filtered[
        (df_filtered['DATA'] >= start_date) &
        (df_filtered['DATA'] <= end_date)
    ]

    df_filtered = df_filtered[
        (df_filtered['FATURAMENTO'] >= faturamento_min) &
        (df_filtered['FATURAMENTO'] <= faturamento_max)
    ]

    # Página Principal
    st.header('Estatísticas (com filtros)')

    ## Faturamento
    faturamento = int(df_filtered['FATURAMENTO'].sum())

    st.subheader(f'Faturamento: R$ {faturamento:,}')

    st.markdown('---')

    ## Top 5 Clientes
    top5_clientes = (
        df_filtered.groupby('CLIENTE')['FATURAMENTO'].sum().sort_values(ascending = False)[:5]
    )

    fig_top5_clientes = px.bar(
        top5_clientes,
        x = top5_clientes.index,
        y = 'FATURAMENTO',
        title = '<b>Top 5 Clientes (Faturamento)</b>',
        template = 'plotly_white'
    )

    ## Top 5 Produtos
    top5_produtos = (
        df_filtered.groupby('PRODUTO')['QUANTIDADE'].sum().sort_values(ascending = False)[:5]
    )

    fig_top5_produtos = px.bar(
        top5_produtos,
        x = top5_produtos.index,
        y = 'QUANTIDADE',
        title = '<b>Top 5 Produtos (Quantidade)</b>',
        template = 'plotly_white'
    )

    left_col, right_col = st.columns(2)
    with left_col:
        st.plotly_chart(fig_top5_clientes)
    with right_col:
        st.plotly_chart(fig_top5_produtos)

    st.markdown('---')

    ## Faturamento por Mês
    faturamento_por_mes = (
        df_filtered.groupby('MÊS', sort=False)['FATURAMENTO'].sum()   
    )

    fig_faturamento_por_mes = px.bar(
        faturamento_por_mes,
        x = faturamento_por_mes.index,
        y = 'FATURAMENTO',
        title = '<b>Faturamento por Mês</b>',
        template = 'plotly_white'
    )

    ## Quantidade por Mês
    quantidade_por_mes = (
        df_filtered.groupby('MÊS', sort=False)['QUANTIDADE'].sum()   
    )

    fig_quantidade_por_mes = px.bar(
        quantidade_por_mes,
        x = quantidade_por_mes.index,
        y = 'QUANTIDADE',
        title = '<b>Quantidade por Mês</b>',
        template = 'plotly_white'
    )

    left_col, right_col = st.columns(2)
    with left_col:
        st.plotly_chart(fig_faturamento_por_mes)
    with right_col:
        st.plotly_chart(fig_quantidade_por_mes)

    st.markdown('---')

    ## Frações do Mercado
    fracao_mercado = (
        df_filtered.groupby('MERCADO')['FATURAMENTO'].sum()   
    )

    fig_fracao_mercado = px.pie(
        fracao_mercado,
        names = fracao_mercado.index,
        values = 'FATURAMENTO',
        title = '<b>Faturamento por Mercado</b>',
        template = 'plotly_white'
    )

    fracao_uf = (
        df_filtered.groupby('UF')['FATURAMENTO'].sum()   
    )

    fig_fracao_uf = px.pie(
        fracao_uf,
        names = fracao_uf.index,
        values = 'FATURAMENTO',
        title = '<b>Faturamento por UF</b>',
        template = 'plotly_white'
    )

    left_col, right_col = st.columns(2)
    with left_col:
        st.plotly_chart(fig_fracao_mercado)
    with right_col:
        st.plotly_chart(fig_fracao_uf)

with tabs[1]:

    st.header('Maiores Faturamentos')

    top30_faturamentos = df[['CLIENTE','FATURAMENTO']].sort_values('FATURAMENTO', ascending = False)
    top30_faturamentos = top30_faturamentos.reset_index(drop=True)
    st.dataframe(top30_faturamentos)

with tabs[2]:

    st.header('Tabela Completa:')

    st.dataframe(df_filtered)

with tabs[3]:
    
    st.header('Baixar Folha do Excel')
    
    file_name = st.text_input('Nome do Arquivo:', 'Banco de Dados PPack', key='Nome do Arquivo')
    
    if st.button('Converter para Excel...'):
        
        if file_name:
            
            # Add .xlsx
            file_name = file_name + '.xlsx'
            
            with st.spinner(text='Convertendo...'):
                
                # Save file
                excel_data = df_to_excel(df)
                st.download_button(
                    label='Baixar Excel',
                    data=excel_data,
                    file_name=file_name,
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )
            
        else:
            st.error('Insira o nome do arquivo.')