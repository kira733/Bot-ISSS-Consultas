import pandas as pd
import requests
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters

# Token del bot de Telegram
TOKEN = '7309205915:AAHIUNf-GxO4Mldp0YDJ5wQME5DkxrZanKw'

# Enlace directo al archivo Excel en Google Drive (descarga directa)
file_url = 'https://docs.google.com/spreadsheets/d/1khXtYhn9daZ03L1hKrgIzvIiwrmYU06d/export?format=xlsx'

# Descargar el archivo desde Google Drive
response = requests.get(file_url)

# Verificar si la descarga fue exitosa
if response.status_code == 200:
    print("Archivo descargado correctamente.")
    # Guardar el archivo temporalmente en local
    with open('archivo.xlsx', 'wb') as f:
        f.write(response.content)
else:
    print(f"Error al descargar el archivo: {response.status_code}")
    exit()

# Leer los datos del archivo Excel con pandas y el motor openpyxl
df = pd.read_excel('archivo.xlsx', engine='openpyxl')

# Función para iniciar el bot
async def start(update: Update, context):
    await update.message.reply_text("Hola, soy el bot de consultas. Puedes buscar por Nombre o DUI de empleados.")

# Función para buscar por nombre o DUI
async def buscar(update: Update, context):
    consulta = update.message.text.strip().lower()

    # Verificar si la consulta es un número (DUI) o texto (nombre)
    if consulta.isdigit():
        # Si es numérico, buscar en la columna 'DUI'
        resultado = df[df['DUI'].astype(str).str.contains(consulta, case=False)]
    else:
        # Si es texto, buscar en la columna 'NOMBRE EMPLEADO' (manejo de mayúsculas y minúsculas, acentos, etc.)
        resultado = df[df['NOMBRE EMPLEADO'].str.lower().str.contains(consulta, case=False)]

    if not resultado.empty:
        # Responder con los resultados filtrados
        for index, row in resultado.iterrows():
            respuesta = (
                f"Patronal: {row['PATRONAL']}\n"
                f"Empresa: {row['NOMBRE']}\n"
                f"Dirección empresa: {row['DIRECCION']}\n"
                f"Teléfono: {row['TELEFONO']}\n"
                f"Período: {row['PERIODO']}\n"
                f"Patronal empleado: {row['PATRONAL EMP']}\n"
                f"Salario: {row['SALARIO']}\n"
                f"Empleado: {row['NOMBRE EMPLEADO']}\n"
                f"Dirección empleado: {row['DIRECCION']}\n"
                f"DUI: {row['DUI']}\n"
                f"NIT: {row['NIT']}"
            )
            await update.message.reply_text(respuesta)
    else:
        await update.message.reply_text(f"No se encontraron resultados para '{consulta}'.")

# Main del bot
async def main():
    # Crear la aplicación del bot
    application = Application.builder().token(TOKEN).build()

    # Añadir los manejadores de comandos
    application.add_handler(CommandHandler("start", start))
    application.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, buscar))

    # Iniciar el bot
    await application.start_polling()
    await application.idle()

if __name__ == '__main__':
    import asyncio
    asyncio.run(main())
