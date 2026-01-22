import pandas as pd
from IPython.display import display

df = pd.read_excel("DataBase.xlsx")
#display (df.head())
# head () mostra as linhas desejas, sejam elas as primeiras.

# Pesquisa pelo usuário
colun1 = df['Nome']
usuario = str(input("Digite o nome do usuário que deseja pesquisar: "))   