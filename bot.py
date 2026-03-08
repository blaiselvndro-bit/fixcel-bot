import os
import pandas as pd
from telegram import Update
from telegram.ext import ApplicationBuilder, MessageHandler, CommandHandler, filters, ContextTypes

TOKEN = os.getenv("BOT_TOKEN")

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Welcome to FIXCEL\n\nUpload an Excel file and I will fix it."
    )

async def handle_file(update: Update, context: ContextTypes.DEFAULT_TYPE):

    file = update.message.document

    if not file.file_name.endswith(".xlsx"):
        await update.message.reply_text("Please upload an Excel file.")
        return

    tg_file = await file.get_file()

    await tg_file.download_to_drive("input.xlsx")

    df = pd.read_excel("input.xlsx")

    df = df.dropna(how="all")
    df = df.drop_duplicates()

    df.to_excel("fixed.xlsx", index=False)

    await update.message.reply_document(document=open("fixed.xlsx","rb"))

app = ApplicationBuilder().token(TOKEN).build()

app.add_handler(CommandHandler("start", start))
app.add_handler(MessageHandler(filters.Document.ALL, handle_file))

app.run_polling()
