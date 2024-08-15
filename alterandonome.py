import pandas as pd

df_modificado = pd.read_excel('planilha_modificada.xlsx')
df_setup = pd.read_excel('SETUP MOVI.xlsx')

print("planilha_modificada:", df_modificado.columns)
print("Colunas em SETUP MOVI:", df_setup.columns)

mapeamento = pd.Series(df_setup['Nome 1'].values, index=df_setup['Contato 1']).to_dict()

df_modificado['Contato'] = df_modificado['Contato'].map(mapeamento).fillna(df_modificado['Contato'])

df_modificado.to_excel('planilha_modificada_com_nomes.xlsx', index=False)
