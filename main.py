import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
from status import calcular_status
from formato import formatar_excel
from atualiza import atualizar_base

def main():
    st.title("Atualização da Base de Dados de Acompanhamento do GA")

    current_date = datetime.now().strftime("%Y_%m_%d")

    uploaded_file_1 = st.file_uploader("Escolha o arquivo Acompanhamento_GA.xlsx", type="xlsx")
    uploaded_file_2 = st.file_uploader("Escolha o arquivo query.xlsx", type="xlsx")

    if uploaded_file_1 is not None and uploaded_file_2 is not None:
        construcao = pd.read_excel(uploaded_file_1, index_col=False, sheet_name='Construção')
        industria = pd.read_excel(uploaded_file_1, index_col=False, sheet_name='Indústria')

        construcao['Data Solicitação'] = pd.to_datetime(construcao['Data Solicitação'])
        industria['Data Solicitação'] = pd.to_datetime(industria['Data Solicitação'])
        construcao['Data Solicitação'] = construcao['Data Solicitação'].dt.date
        industria['Data Solicitação'] = industria['Data Solicitação'].dt.date

        consulta = pd.read_excel(uploaded_file_2, index_col=False)
        consulta = consulta[['Solicitação n°', 'Status da Atividade', 'Solicitante', 'Executor', 
                             'Tipo de solicitação', 'JOB/Serviço', 'Prazo p/retorno', 'Solicitado em', 'Solicitado por']]

        consulta = consulta.rename(columns={
            'Solicitação n°': 'Solicitação',
            'Status da Atividade': 'Status do GA',
            'Solicitante': 'Área',
            'Executor': 'Executor',
            'Tipo de solicitação': 'Tipo de Solicitação',
            'JOB/Serviço': 'JOB',
            'Prazo p/retorno': 'Data Retorno',
            'Solicitado em': 'Data Solicitação',
            'Solicitado por': 'Solicitante'
        })

        consulta['Data Solicitação'] = pd.to_datetime(consulta['Data Solicitação'])
        consulta['Data Solicitação'] = consulta['Data Solicitação'].dt.date

        consulta_construcao = consulta[consulta['Área'].str.contains('Construção', case=False, na=False)]
        consulta_industria = consulta[consulta['Área'].str.contains('Indústria', case=False, na=False)]

        construcao_atualizada = atualizar_base(construcao, consulta_construcao)
        industria_atualizada = atualizar_base(industria, consulta_industria)

        construcao_atualizada['Status da Demanda'] = construcao_atualizada.apply(calcular_status, axis=1)
        industria_atualizada['Status da Demanda'] = industria_atualizada.apply(calcular_status, axis=1)

        construcao_atualizada['Status da Demanda'] = construcao_atualizada.apply(lambda row: 'Finalizada' if row['Status do GA'] in ['Concluída', 'Cancelada'] else row['Status da Demanda'], axis=1)
        industria_atualizada['Status da Demanda'] = industria_atualizada.apply(lambda row: 'Finalizada' if row['Status do GA'] in ['Concluída', 'Cancelada'] else row['Status da Demanda'], axis=1)

        construcao_atualizada['Data Retorno'] = construcao_atualizada['Data Retorno'].dt.date
        industria_atualizada['Data Retorno'] = industria_atualizada['Data Retorno'].dt.date

        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            formatar_excel(writer, 'Construção', construcao_atualizada)
            formatar_excel(writer, 'Indústria', industria_atualizada)
        output.seek(0)

        st.success("A base de dados foi atualizada com sucesso!")
        st.download_button(label="Baixar Acompanhamento_GA.xlsx", data=output, file_name=f'Acompanhamento_GA_{current_date}.xlsx')
    
    else:
        st.warning('Por favor, carregue ambos os arquivos para continuar.')

if __name__ == "__main__":
    main()
