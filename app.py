import base64
import os
import platform
import smtplib
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import numpy as np
import pandas as pd
import streamlit as st
import teradata
from sqlalchemy import create_engine

if platform.system() == "Windows":
    import win32com.client

# Configurações da página
st.set_page_config(
    page_title="GEO - Indicadores Operacionais",
    page_icon="🧙‍♂️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Carregar dados
data = pd.read_excel('Acesso.xlsx')

# Função para download do CSV
def download_csv(df, loja1, loja2):
    # Selecione as colunas necessárias
    df = df[['NOM_DEPTO', 'NOM_PLU', 'COD_PLU', loja1, loja2, '∆ Delta']]

    # Converta o DataFrame para CSV e depois para base64
    csv = df.to_csv(index=False)
    b64 = base64.b64encode(csv.encode()).decode()

    # Crie um link de download
    href = f'<a href="data:file/csv;base64,{b64}" download="comparacao_sortimento.csv">Clique aqui para Baixar CSV</a>'
    st.markdown(href, unsafe_allow_html=True)


# Inicializar o estado da sessão, se necessário
if 'logged_in' not in st.session_state:
    st.session_state.logged_in = False

if not st.session_state.logged_in:
    user = st.text_input("Matricula")
    password = st.text_input("Senha", type='password')

    if st.button('Login'):
        try:
            user = int(user)
            password = int(password)

            if (data['Matricula'] == user).any() and (data['Senha'] == password).any():
                st.success('Você está logado.')
                st.session_state.logged_in = True
                st.rerun()
            else:
                st.error(
                    'As credenciais estão incorretas. Tente novamente os 3 primeiros digitos de sua Matricula.')
        except ValueError:
            st.error('A matrícula e a senha devem ser números inteiros.')
else:
    st.sidebar.title("Menu")
    option = st.sidebar.selectbox(
        'Selecione uma opção',
        ('HomePage', 'Comparação de Lojas', 'Comparação de Sortimento',
         'DDP D0', 'Store Visit', 'Quadro de Funcionarios', 'Farol Operacional')
    )

    def find_similar_store(row, prox, columns, k=1):
        prox['distance'] = np.sqrt(
            (prox[columns[0]] - row[columns[0]])**2 + (prox[columns[1]] - row[columns[1]])**2)
        prox = prox[prox['CodLoja'] != row['CodLoja']]
        nearest_points = prox.nsmallest(k, 'distance')
        return nearest_points['CodLoja'].values if len(nearest_points) >= k else [None]*k

    def display_prox(prox, loja, columns_to_display):
        if 'CodLoja' not in prox.columns:
            st.error("A coluna 'CodLoja' não está presente no DataFrame.")
            return

        prox_selected = prox[prox['CodLoja'] == loja]
        if prox_selected.empty:
            st.error(f"Nenhuma loja encontrada com o código {loja}.")
            return

        lojas_comparacao = prox_selected[[
            '1°_Comparacao', '2°_Comparacao', '3°_Comparacao']].values[0]
        prox_comparacao = prox[prox['CodLoja'].isin(lojas_comparacao)]

        # Exibir apenas as colunas desejadas
        st.write(prox_comparacao[columns_to_display])

    def consultar_teradata(loja1, loja2):
        caminho_arquivo = 'Sortimento.xlsx'
        df = pd.read_excel(caminho_arquivo)
        lojas = [loja1, loja2]
        df_filtrado = df[df['COD_LOJA'].isin(lojas)]
        return df_filtrado

    if option == 'HomePage':
        with st.container():
            st.title("GEO - Indicadores Operacionais")
            st.write("Bem-vindo ao Sistema de Comparação de Lojas")
            st.write(
                "Para verificar nossos DashBoards [Clique aqui](https://app.powerbi.com/groups/me/apps/9190d269-d305-4cb8-bbb1-a63a633498a6/reports/9a0e6a1d-8b2e-4cf2-a544-54fc20d8f97b/ReportSection66f47a334e97cc09a9ab?experience=power-bi)")

    elif option == 'Comparação de Lojas':
        prox = pd.read_excel('Comparacao_Lojas.xlsx')

        loja_input = st.text_input("Digite o número da loja")
        if loja_input:
            loja = int(loja_input)

            prox['Latitude'] = pd.to_numeric(prox['Latitude'], errors='coerce')
            prox['Longitude'] = pd.to_numeric(
                prox['Longitude'], errors='coerce')
            prox['QTD_TAMANHO_AREA_VENDA'] = pd.to_numeric(
                prox['QTD_TAMANHO_AREA_VENDA'], errors='coerce')
            prox['MEDIA_DIARIA_VENDA'] = pd.to_numeric(
                prox['MEDIA_DIARIA_VENDA'], errors='coerce')

            prox['QTD_TAMANHO_AREA_VENDA'].fillna(0, inplace=True)
            prox['MEDIA_DIARIA_VENDA'].fillna(0, inplace=True)

            results = prox.apply(find_similar_store, axis=1, prox=prox, columns=[
                                 'MEDIA_DIARIA_VENDA', 'QTD_TAMANHO_AREA_VENDA'], k=3)
            prox['1°_Comparacao'], prox['2°_Comparacao'], prox['3°_Comparacao'] = zip(
                *results)

            if 'distance' in prox.columns:
                prox = prox.drop(columns=['distance'])

            columns_to_display = ['Loja', 'Formato', 'MicroRegiaoFinal',
                                  'QTD_TAMANHO_AREA_VENDA', 'MEDIA_DIARIA_VENDA']
            display_prox(prox, loja, columns_to_display)

    elif option == 'Comparação de Sortimento':
        loja1_input = st.text_input("Digite o código da Loja 1")
        loja2_input = st.text_input("Digite o código da Loja 2")
        if loja1_input and loja2_input:
            loja1 = int(loja1_input)
            loja2 = int(loja2_input)
            df_filtrado = consultar_teradata(loja1, loja2)

            pivot_df = df_filtrado.pivot_table(index=['NOM_DEPTO'],
                                    columns='COD_LOJA', values='COD_PLU', aggfunc='count', fill_value=0)
            pivot_df = pivot_df.dropna(how='all').fillna(0)
            pivot_df.columns.name = None
            pivot_df = pivot_df.reset_index()

            if loja1 in pivot_df.columns and loja2 in pivot_df.columns:
                pivot_df['∆ Delta'] = pivot_df[loja1] - pivot_df[loja2]
            else:
                st.write(
                    f"As lojas '{loja1}' e '{loja2}' não existem no DataFrame.")

            # Calcular a linha de total
            total_row = {
                'NOM_DEPTO': 'Total',
                loja1: pivot_df[loja1].sum(),
                loja2: pivot_df[loja2].sum(),
                '∆ Delta': pivot_df['∆ Delta'].abs().sum()
            }

            # Adicionar a linha "Total" ao DataFrame
            total_row_df = pd.DataFrame(total_row, index=[0])
            pivot_df = pd.concat([pivot_df, total_row_df], ignore_index=True)

            st.write(pivot_df)

            # Chame a função quando o botão 'Download CSV' for clicado
            if st.button('Download CSV'):
                download_csv(df_filtrado, loja1, loja2)

    elif option == 'DDP D0':
        caminho_arquivo = 'DDP_D0.xlsx'
        df = pd.read_excel(caminho_arquivo)

        loja_input = st.text_input("Digite o código da loja")
        if loja_input:
            loja = int(loja_input)
            df_filtrado = df[df['cod_loja'] == loja]
            columns_to_display = [
                'cod_loja', 'cod_plu', 'Nom_Prod', 'dta_analise', 'horaatual', 'horaultvda']
            st.write(df_filtrado[columns_to_display])

            if not df_filtrado.empty:
                csv = df_filtrado.to_csv(index=False)
                b64 = base64.b64encode(csv.encode()).decode()
                href = f'<a href="data:file/csv;base64,{b64}" download="loja_{loja}.csv">Clique aqui para Baixar CSV</a>'
                st.markdown(href, unsafe_allow_html=True)

    elif option == 'Store Visit':

        loja_input = st.text_input("Digite o Código da Loja")
        email_input = st.text_input("Digite o Email")

        if loja_input and email_input:
            loja = int(loja_input)
            email = email_input

            if st.button('Solicitar'):
                # Configurações do servidor SMTP
                smtp_server = 'smtp.outlook.com'
                smtp_port = 587
                smtp_user = 'mateus.santos2@gpabr.com'
                smtp_password = 'Gpa@982110764'

                msg = MIMEMultipart()
                msg['From'] = smtp_user
                msg['To'] = smtp_user
                msg['Subject'] = 'Solicitação de Store visit'

                body = f'O usuário solicitou um Store Visit para a loja {loja} com o e-mail {email}.'
                msg.attach(MIMEText(body, 'plain'))

                try:
                    with smtplib.SMTP(smtp_server, smtp_port) as server:
                        server.starttls()
                        server.login(smtp_user, smtp_password)
                        server.sendmail(smtp_user, smtp_user, msg.as_string())

                    st.success('Solicitação enviada com sucesso!')
                except Exception as e:
                    st.error(f"Erro ao enviar solicitação: {str(e)}")

    elif option == 'Quadro de Funcionarios':
        # Carregar dados do quadro de funcionários a partir de um arquivo Excel
        dfquadro = pd.read_excel('Quadro_Funcionarios.xlsx')

        loja_input = st.text_input("Digite o número da loja")
        if loja_input:
            loja = int(loja_input)

            # Filtrar os dados para a loja específica
            df_loja = dfquadro[dfquadro['COD_LOJA'] == loja]

            if df_loja.empty:
                st.error(f"Nenhuma loja encontrada com o código {loja}.")
            else:
                # Calcular as novas colunas
                df_loja['<> DELTA'] = df_loja['COLABS_SUGERIDOS'] - \
                    df_loja['COLABS_ATIVOS']

                # Calcular % ORC x ATIVO e garantir que não ultrapasse 100%
                df_loja['% ORC x ATIVO'] = (
                    df_loja['COLABS_ATIVOS'] / df_loja['COLABS_SUGERIDOS']) * 100
                df_loja['% ORC x ATIVO'] = df_loja['% ORC x ATIVO'].apply(
                    lambda x: min(x, 100)).round(2)

                # Formatando como porcentagem com 2 casas decimais
                df_loja['% ORC x ATIVO'] = df_loja['% ORC x ATIVO'].astype(
                    str) + '%'

                # Calcular totais
                total_row = {
                    'COD_LOJA': 'Total',
                    'Loja': '',
                    'SETORES': '',
                    'COLABS_SUGERIDOS': df_loja['COLABS_SUGERIDOS'].sum(),
                    'COLABS_ATIVOS': df_loja['COLABS_ATIVOS'].sum(),
                    '<> DELTA': df_loja['<> DELTA'].sum(),
                    '% ORC x ATIVO': f"{min((df_loja['COLABS_ATIVOS'].sum() / df_loja['COLABS_SUGERIDOS'].sum()) * 100, 100):.2f}%"
                }

                # Adicionar a linha "Total" ao DataFrame
                total_row_df = pd.DataFrame(total_row, index=[0])
                df_loja = pd.concat([df_loja, total_row_df], ignore_index=True)

                # Selecionar as colunas a serem exibidas
                columns_to_display = ['COD_LOJA', 'Loja', 'SETORES',
                                      'COLABS_SUGERIDOS', 'COLABS_ATIVOS', '<> DELTA', '% ORC x ATIVO']
                st.write(df_loja[columns_to_display])

    elif option == 'Farol Operacional':
        loja_option = st.selectbox(
            'Selecione uma Bandeira',
            ('PÃO DE AÇÚCAR', 'MERCADO EXTRA')
        )

        if loja_option == 'PÃO DE AÇÚCAR':
            caminho_do_arquivo = "GPA_FAROL_Painel_PA_v4.pdf"
            with open(caminho_do_arquivo, "rb") as file:
                st.download_button(
                    label="Baixar PDF de PÃO DE AÇÚCAR",
                    data=file,
                    file_name="PÃO_DE_AÇÚCAR.pdf",
                    mime="application/pdf"
                )
        elif loja_option == 'MERCADO EXTRA':
            caminho_do_arquivo = "GPA_FAROL_Painel_ME_v4.pdf"
            with open(caminho_do_arquivo, "rb") as file:
                st.download_button(
                    label="Baixar PDF de MERCADO EXTRA",
                    data=file,
                    file_name="MERCADO_EXTRA.pdf",
                    mime="application/pdf"
                )
