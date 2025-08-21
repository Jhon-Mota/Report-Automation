import pandas as pd

df = pd.read_excel("")

df.dropna(subset = ["Id Loja", "Id Produto"])

sales = df.groupby("Id Loja")["Qtd. Vendida"].sum().sort_values(ascending = False).head()

print(sales)