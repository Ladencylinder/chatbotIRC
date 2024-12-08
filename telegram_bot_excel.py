import openpyxl
import telebot
import time
from threading import Timer

# Token del bot de Telegram
TOKEN = '8117026808:AAFxaLJIvuebS07K2TLUUCgQsaK-qCoNdw0'

# Ruta del archivo de Excel
EXCEL_PATH = EXCEL_PATH = r"INFLACION 10 AÑOS prueba chatbot.xlsx"

# Crear una instancia del bot
bot = telebot.TeleBot(TOKEN)

# Variable para manejar temporizadores
active_sessions = {}

# Función para leer datos del archivo de Excel
def leer_excel():
    try:
        wb = openpyxl.load_workbook(EXCEL_PATH)
        sheet = wb.active
        datos = []
        for row in sheet.iter_rows(values_only=True):
            datos.append(row)
        return datos
    except Exception as e:
        return str(e)

# Función para cerrar sesión tras 1 minuto de inactividad
def cerrar_interaccion(chat_id):
    bot.send_message(chat_id, "Sesión cerrada por inactividad. Usa /start para comenzar de nuevo.")
    if chat_id in active_sessions:
        del active_sessions[chat_id]

# Mensaje de inicio
@bot.message_handler(commands=['start'])
def send_welcome(message):
    chat_id = message.chat.id
    bot.reply_to(message, "Hola, soy tu chatbot. Para realizar una consulta de los precios de los productos y analizar la inflación, envíame el nombre del producto seguido del año solo puedes consultar del 2013 al 2023. Por ejemplo: 'Arroz 2020'.\n"
                              "Lista de artículos disponibles para consulta:\n"
                              "Arroz\n"
                              "Carne de res\n"
                              "Vísceras de res\n"
                              "Frijol procesado\n"
                              "Café soluble\n"
                              "Café tostado\n"
                              "Chocolate y productos de confitería\n"
                              "Chocolate líquido y para preparar bebida\n"
                              "Reproductores de video\n"
                              "Detergentes\n"
                              "Analgésicos\n"
                              "Expectorantes y descongestivos\n"
                              "Otros medicamentos\n"
                              "Material de curación\n"
                              "Jabón de tocador\n"
                              "Autobús urbano\n"
                              "Motocicletas\n"
                              "Bicicletas.")
    # Reiniciar temporizador
    if chat_id in active_sessions:
        active_sessions[chat_id].cancel()
    active_sessions[chat_id] = Timer(60, cerrar_interaccion, [chat_id])
    active_sessions[chat_id].start()

# Manejo de mensajes
@bot.message_handler(func=lambda message: True)
def echo_all(message):
    try:
        chat_id = message.chat.id
        # Reiniciar temporizador
        if chat_id in active_sessions:
            active_sessions[chat_id].cancel()
        active_sessions[chat_id] = Timer(60, cerrar_interaccion, [chat_id])
        active_sessions[chat_id].start()

        query = message.text.split()
        if len(query) < 2:
            bot.reply_to(message, "Por favor, envía tu consulta en el formato: 'Producto Año'. Por ejemplo: 'Arroz 2020'.")
            return

        producto = " ".join(query[:-1]).lower().strip()
        year = query[-1].strip()

        if not year.isdigit():
            bot.reply_to(message, "El año debe ser un número. Por favor, envía tu consulta correctamente.")
            return

        year = int(year)

        # Validar rango de años
        if year < 2013 or year > 2023:
            bot.reply_to(message, "El año debe estar entre 2013 y 2023. Por favor, intenta nuevamente.")
            return

        datos = leer_excel()

        if isinstance(datos, str):  # Si hubo un error leyendo el archivo
            bot.reply_to(message, f"Error leyendo el archivo de Excel: {datos}")
            return

        # Buscar coincidencias
        encabezados = datos[0]
        filas = datos[1:]  # Omitir encabezados
        resultados = [
            fila for fila in filas
            if fila[0] == year and fila[2].strip().lower() == producto
        ]

        # Dividir resultados en bloques y enviar con un retraso
        if resultados:
            bloque_tamaño = 5  # Máximo de resultados por mensaje
            for i in range(0, len(resultados), bloque_tamaño):
                bloque = resultados[i:i+bloque_tamaño]
                respuesta = "\n".join([
                    " | ".join(f"{encabezados[j]}: {round(float(celda), 2) if isinstance(celda, float) else celda}" for j, celda in enumerate(fila))
                    for fila in bloque
                ])
                bot.reply_to(message, f"Resultados:\n{respuesta}")
                time.sleep(1)  # Añade un retraso de 1 segundo entre mensajes
        else:
            bot.reply_to(message, f"No se encontraron resultados para el producto '{producto}' en el año {year}.")

        # Finalizar con mensaje de reinicio
        bot.reply_to(message, "Consulta finalizada. Usa /start para comenzar nuevamente.")

    except Exception as e:
        bot.reply_to(message, f"Ocurrió un error procesando tu consulta: {e}")

# Iniciar el bot
print("Bot en funcionamiento...")
bot.polling()
