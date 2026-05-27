import mysql.connector
import pandas as pd
from datetime import datetime
import smtplib
from email.message import EmailMessage
from dotenv import load_dotenv
import os

load_dotenv()

conexion = mysql.connector.connect(
    host= os.getenv("HOST"),
    user=os.getenv("USER"),
    password=os.getenv("PASSWORD"),
    database=os.getenv("NAME_DB")
)


query = """
SELECT *
FROM sales
WHERE date = CURDATE();
"""

df = pd.read_sql(query, conexion)


fecha = datetime.now().strftime("%Y-%m-%d")

archivo_excel = f"ventas_{fecha}.xlsx"

df.to_excel(archivo_excel, index=False)

print("Excel generado correctamente")

EMAIL = os.getenv("EMAIL")
PASSWORD = os.getenv("EMAIL_PASSWORD")

DESTINO = "destino@gmail.com"

msg = EmailMessage()

msg["Subject"] = f"Reporte de ventas {fecha}"
msg["From"] = EMAIL
msg["To"] = DESTINO

msg.set_content("Adjunto reporte automático de ventas del día.")


with open(archivo_excel, "rb") as f:
    file_data = f.read()
    file_name = f.name

msg.add_attachment(
    file_data,
    maintype="application",
    subtype="octet-stream",
    filename=file_name
)


with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
    smtp.login(EMAIL, PASSWORD)
    smtp.send_message(msg)

print("Correo enviado correctamente")