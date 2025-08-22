import pandas as pd
import seaborn as sea
import matplotlib.pyplot as plt

df = pd.read_excel("")

df.dropna(subset = ["Id Loja", "Id Produto"])

sales = df.groupby("Id Loja")["Qtd. Vendida"].sum().head()

sales_df = sales.reset_index()


sea.set_theme(style="darkgrid")
sea.barplot(data=sales_df, x="Id Loja", y="Qtd. Vendida")

plt.show()

#print(sales)