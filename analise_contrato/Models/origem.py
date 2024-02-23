import pandas as pd

# Função para gerar os dias 01 de cada mês entre duas datas
def gerar_dias_01_entre_datas(data_inicio, data_fim):
    datas_01 = pd.date_range(start=data_inicio, end=data_fim, freq='MS')
    return datas_01

# Função para formatar a data no estilo brasileiro (DD/MM/AAAA)
def formatar_data_brasileira(data):
    return data.strftime('%d/%m/%Y 00')

# Ler o arquivo Excel
file_path = ''
with pd.ExcelWriter(file_path, mode='a', engine='openpyxl') as writer:
    df = pd.read_excel(file_path, sheet_name='Extrato', usecols=range(11))

    # Selecionar as colunas específicas (id e datas)
    df = df[['CÓDIGO', 'INÍCIO DE VIGÊNCIA', 'FIM DE VIGÊNCIA']]

    # Inicializar um dicionário para armazenar os DataFrames de cada código
    codigo_dfs = {}

    for codigo, group in df.groupby('CÓDIGO'):
        datas_01_brasil = []  # Inicializar a lista de datas brasileiras

        # Iterar sobre cada linha do grupo
        for _, row in group.iterrows():
            data_inicio = row['INÍCIO DE VIGÊNCIA']
            data_fim = row['FIM DE VIGÊNCIA']

            # Gerar as datas do dia 01 entre as datas de início e fim
            datas_01 = gerar_dias_01_entre_datas(data_inicio, data_fim)

            # Formatar e adicionar as datas do dia 01 à lista, evitando duplicatas
            datas_01_brasil.extend([formatar_data_brasileira(data) for data in datas_01 if
                                    formatar_data_brasileira(data) not in datas_01_brasil])

        # Criar um DataFrame para este grupo de código
        df_temp = pd.DataFrame({'Datas Validação': datas_01_brasil, 'Condição': 'X', 'CÓDIGO': codigo})

        # Adicionar o DataFrame deste código ao dicionário
        codigo_dfs[codigo] = df_temp

    # Concatenar todos os DataFrames do dicionário em um único DataFrame
    novo_df = pd.concat(codigo_dfs.values(), ignore_index=True)

    # Reordenar as colunas
    novo_df = novo_df[['CÓDIGO', 'Datas Validação', 'Condição']]

    # Colocar em branco as células na coluna 'CÓDIGO' que têm o mesmo código repetido
    novo_df.loc[novo_df['CÓDIGO'].duplicated(), 'CÓDIGO'] = ''

    # Centralizar o conteúdo das colunas
    novo_df_styled = novo_df.style.set_properties(**{'text-align': 'center'})

    # Adicionar a nova planilha ao arquivo Excel original
    novo_df.to_excel(writer, sheet_name='VALIDAÇÃO', index=False)
