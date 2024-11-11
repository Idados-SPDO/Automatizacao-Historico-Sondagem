import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(
    page_title="APP Atualização de histórico - Sondagem",
    layout="wide"
)

st.header("Ferramenta de Atualização de histórico - Sondagem", divider="rainbow")

st.write('')
st.write('')

def baixar_modelo(dataframes, arquivo):
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for sheet_name, df in dataframes.items():
            df.to_excel(writer, index=False, sheet_name=sheet_name)
    buffer.seek(0)  # Volte ao início do buffer
    return st.download_button(label="Baixar Modelo de Input", data=buffer, file_name=f"{arquivo}.xlsx")

modelo = pd.read_excel('Modelo.xlsx', sheet_name=None)

with st.expander("Dados necessários"):
    st.markdown('''
    Para carregar os dados relacionados às **Sondagens**, é necessário que o arquivo de input esteja igual ao apresentado abaixo:
    ''')
    baixar_modelo(modelo, "Modelo")

st.write('')

tipo_sondagem = st.selectbox("**Sondagem**", options=["Comércio", "Construção", "Consumidor", "Indústria", "SEBRAE", "Serviços"])

# Crio um container para 'encapsular' itens dentro de uma parte da tela.
container = st.container()

# Solicito que o usuário selecione uma data e imprimo na tela a data selecionada.
container.subheader('Defina a data de referência e selecione o arquivo para atualizar o histórico:', divider="rainbow")
data_referencia = container.date_input("", value=None, format="DD/MM/YYYY")
arquivo_atualizacao = st.file_uploader("arquivo_atualizacao", type="xlsx", label_visibility="hidden")

if data_referencia:
    # Alterar o dia para o primeiro dia do mês
    data_referencia = data_referencia.replace(day=1)

def load_mailing(content_file=None):
    if content_file is not None:
        content = pd.read_excel(content_file, sheet_name="Mailing")

        return content

def load_placar(content_file=None):
    if content_file is not None:
        content = pd.read_excel(content_file, sheet_name="Placar")

        return content

def load_telefone(content_file=None):
    if content_file is not None:
        content = pd.read_excel(content_file, sheet_name="Telefone")

        return content

def load_meta(content_file=None):
    if content_file is not None:
        content = pd.read_excel(content_file, sheet_name="Meta")

        return content


if arquivo_atualizacao is not None:
    if tipo_sondagem != 'SEBRAE':
        df_mailing = load_mailing(arquivo_atualizacao)
        df_mailing['Data de Referência'] = data_referencia
        df_mailing_historico = pd.read_excel(f'\\fgvfsbi\ibre-sci-sapc\Otimização e Automatização/5. Projetos/22. Produção dos Pesquisadores/1. Dados/Histórico/historico_{tipo_sondagem}.xlsx', sheet_name = "Mailing", dtype={'CNAE Princ': str}, parse_dates=['Data de Referência'])
        df_mailing_historico['Data de Referência'] = df_mailing_historico['Data de Referência'].dt.date
        df_mailing_atual = pd.concat([df_mailing_historico, df_mailing])
        df_mailing_atual.drop_duplicates(inplace = True)
        df_mailing_atual.reset_index(drop = True, inplace = True)

    df_placar = load_placar(arquivo_atualizacao)
    df_placar['Data de Referência'] = data_referencia
    df_placar_historico = pd.read_excel(f'\\fgvfsbi\ibre-sci-sapc\Otimização e Automatização/5. Projetos/22. Produção dos Pesquisadores/1. Dados/Histórico/historico_{tipo_sondagem}.xlsx', sheet_name = "Placar")
    df_placar_historico['Data de Referência'] = df_placar_historico['Data de Referência'].dt.date
    df_placar_atual = pd.concat([df_placar_historico, df_placar])
    df_placar_atual.drop_duplicates(inplace = True)
    df_placar_atual.reset_index(drop = True, inplace = True)
    
    df_telefone = load_telefone(arquivo_atualizacao)
    df_telefone['Data'] = df_telefone['Data'].dt.date
    df_telefone_historico = pd.read_excel(f'\\fgvfsbi\ibre-sci-sapc\Otimização e Automatização/5. Projetos/22. Produção dos Pesquisadores/1. Dados/Histórico/historico_{tipo_sondagem}.xlsx', sheet_name = "Telefone", parse_dates=['Data'])
    df_telefone_historico['Data'] = df_telefone_historico['Data'].dt.date
    df_telefone_atual = pd.concat([df_telefone_historico, df_telefone])
    df_telefone_atual.drop_duplicates(inplace = True)
    df_telefone_atual.reset_index(drop = True, inplace = True)
    
    df_meta = load_meta(arquivo_atualizacao)
    df_meta['Data de Referência'] = data_referencia
    df_meta_historico = pd.read_excel(f'\\fgvfsbi\ibre-sci-sapc\Otimização e Automatização/5. Projetos/22. Produção dos Pesquisadores/1. Dados/Histórico/historico_{tipo_sondagem}.xlsx', sheet_name = "Meta", parse_dates=['Data de Referência'])
    df_meta_historico['Data de Referência'] = df_meta_historico['Data de Referência'].dt.date
    df_meta_atual = pd.concat([df_meta_historico, df_meta])
    df_meta_atual.drop_duplicates(inplace = True)
    df_meta_atual.reset_index(drop = True, inplace = True)

    st.subheader('', divider="rainbow")
    st.write('')
    st.subheader('Confira as bases abaixo e caso esteja tudo correto, atualize o histórico.')
    if tipo_sondagem != 'SEBRAE':
        st.write('')
        st.write('**Mailing:**')
        st.write(df_mailing_atual)
    st.write('')
    st.write('**Placar:**')
    st.write(df_placar_atual)
    st.write('')
    st.write('**Telefone:**')
    st.write(df_telefone_atual)
    st.write('')
    st.write('**Meta:**')
    st.write(df_meta_atual)
    st.write('')

    # Cria o botão
    if st.button('Clique aqui para atualizar o histórico'):
        with pd.ExcelWriter(f'\\fgvfsbi\ibre-sci-sapc\Otimização e Automatização/5. Projetos/22. Produção dos Pesquisadores/1. Dados/Histórico/historico_{tipo_sondagem}.xlsx') as writer:
            if tipo_sondagem != 'SEBRAE':
                df_mailing_atual.to_excel(writer, sheet_name='Mailing', index=False)
            df_placar_atual.to_excel(writer, sheet_name='Placar', index=False)
            df_telefone_atual.to_excel(writer, sheet_name='Telefone', index=False)
            df_meta_atual.to_excel(writer, sheet_name='Meta', index=False)
            st.warning(f'Histórico **{tipo_sondagem}** atualizado com sucesso :)')
