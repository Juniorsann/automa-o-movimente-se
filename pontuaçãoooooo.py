import pandas as pd

nome_arquivo = 'mensagens_encontradas_13_08.xlsx'
df = pd.read_excel(nome_arquivo)

df['Data de Envio'] = pd.to_datetime(df['Data de Envio'], dayfirst=True)

df['Dia da Semana'] = df['Data de Envio'].dt.day_name()

df['Pontuação'] = 10 

pontuacao_diaria = df.groupby(['Contato', df['Data de Envio'].dt.date, 'Dia da Semana'])['Pontuação'].sum().reset_index()
pontuacao_diaria.rename(columns={'Data de Envio': 'Data'}, inplace=True)
pontuacao_diaria['Data'] = pd.to_datetime(pontuacao_diaria['Data'])

contatos = df['Contato'].unique()
datas = pd.date_range(start=pontuacao_diaria['Data'].min(), end=pontuacao_diaria['Data'].max(), freq='D')
todas_combinacoes = pd.MultiIndex.from_product([contatos, datas], names=['Contato', 'Data']).to_frame(index=False)
todas_combinacoes['Data'] = pd.to_datetime(todas_combinacoes['Data'])

pontuacao_completa = todas_combinacoes.merge(pontuacao_diaria, on=['Contato', 'Data'], how='left').fillna(0)
pontuacao_completa['Dia da Semana'] = pontuacao_completa['Data'].dt.day_name()

pivot = pontuacao_completa.pivot_table(index='Contato', columns='Data', values='Pontuação', fill_value=0).reset_index()
pivot.columns = [pivot.columns[0]] + sorted(pivot.columns[1:], key=lambda x: pd.to_datetime(x))
pivot.columns = [pivot.columns[0]] + [f"{col} ({pd.to_datetime(col).day_name()})" for col in pivot.columns[1:]]

pivot['Soma Semanal'] = pivot.iloc[:, 1:].sum(axis=1)

nome_arquivo_tratado = 'pontuacao_diaria_com_dias_semana_e_soma_completa.xlsx'
pivot.to_excel(nome_arquivo_tratado, index=False)

print(f'Pontuação diária com dias da semana e soma semanal salva em {nome_arquivo_tratado}')





