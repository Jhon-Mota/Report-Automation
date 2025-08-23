import pandas as pd
import seaborn as sea
import matplotlib.pyplot as plt
import smtplib
from email.message import EmailMessage
import mimetypes

df = pd.read_excel("./Base Vendas - 2022.xlsx")

df.dropna(subset = ["Id Loja", "Id Produto"])

sales = df.groupby("Id Loja")["Qtd. Vendida"].sum().head()
sales_sorted = sales.sort_values(ascending = False)

sales_df = sales_sorted.reset_index()

sea.set_theme(style="darkgrid")
plot = sea.barplot(data=sales_df, x="Id Loja", y="Qtd. Vendida")

plt.savefig("barplot.png", format="png")

sales_df.to_excel("Top 5 sales.xlsx")

# Email
recipient = ""
sender = ""
emailSubject = "Relatório"
body = """
Relatório com os dados de vendas.

Att, bot de e-mail.
"""
password = ""
arquive = "./Top 5 sales.xlsx"
plot = "./barplot.png"

msg = EmailMessage()
msg['From'] = recipient
msg['To'] = sender
msg['Subject'] = emailSubject
msg.set_content(body)

mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
mime_type, mime_subtype = mime_type.split('/')

mimeimg_type = "image/png"
mimeimg_type, mimeimg_subtype = mimeimg_type.split('/')


try:
    with open(arquive, "rb") as file:
        msg.add_attachment(file.read(), maintype=mime_type, subtype=mime_subtype, filename=arquive)

except Exception as e:
    print("An error occurred.")
    print(e)

try:
    with open(plot, "rb") as image:
        msg.add_attachment(image.read(), maintype=mimeimg_type, subtype=mimeimg_subtype, filename=plot)

except Exception as e:
    print("An error occurred.")
    print(e)


try:
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as email:
        email.login(recipient, password)
        email.send_message(msg)

except Exception as e:
    print("An error has occurred. Please, try again later.")
    print(e)