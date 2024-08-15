import pandas as pd
import re

df = pd.read_excel('pontuacao_diaria_com_dias_semana_e_soma_completa.xlsx')

print("Colunas disponíveis:", df.columns)

print("Exemplos de contatos antes da modificação:", df['Contato'].head())

def clean_contact(contact):
    if isinstance(contact, str):
        
        contact = re.sub(r'^\s*\+55', '', contact)
        return contact.strip()
    return contact

df['Contato'] = df['Contato'].apply(clean_contact)

print("Exemplos de contatos após a modificação:", df['Contato'].head())

df.to_excel('planilha_modificada.xlsx', index=False)
