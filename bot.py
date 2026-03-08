import os
import pandas as pd

from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    ApplicationBuilder,
    MessageHandler,
    CommandHandler,
    CallbackQueryHandler,
    filters,
    ContextTypes
)

from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

TOKEN = os.getenv("BOT_TOKEN")

# store uploaded files temporarily
user_files = {}


# START COMMAND

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):

    text = """
🎨 *Welcome to FIXCEL*

Send an Excel file and I will format it beautifully.

Features
• Custom header color
• Alternating rows
• Wrap text
• Clean borders
• Smart alignment

Free Plan
2 files per month

Premium
Unlimited formatting
$6/month
"""

    await update.message.reply_text(text, parse_mode="Markdown")


# FILE UPLOAD

async def handle_file(update: Update, context: ContextTypes.DEFAULT_TYPE):

    user = update.message.from_user.id
    file = update.message.document

    tg_file = await file.get_file()

    path = f"{user}_input.xlsx"

    await tg_file.download_to_drive(path)

    user_files[user] = path

    keyboard = [
        [InlineKeyboardButton("Use Default Color (#1D6F42)", callback_data="default")]
    ]

    reply_markup = InlineKeyboardMarkup(keyboard)

    await update.message.reply_text(
        "🎨 Send a HEX color for the header.\nExample:\n#FF5733\n\nOr choose default:",
        reply_markup=reply_markup
    )


# DEFAULT COLOR BUTTON

async def default_color(update: Update, context: ContextTypes.DEFAULT_TYPE):

    query = update.callback_query
    await query.answer()

    user = query.from_user.id

    if user not in user_files:
        await query.message.reply_text("Please upload an Excel file first.")
        return

    file = user_files[user]

    result = format_excel(file, "#1D6F42")

    await query.message.reply_document(document=open(result, "rb"))


# RECEIVE HEX

async def receive_hex(update: Update, context: ContextTypes.DEFAULT_TYPE):

    user = update.message.from_user.id

    if user not in user_files:
        return

    hex_color = update.message.text.strip()

    if not hex_color.startswith("#") or len(hex_color) != 7:

        await update.message.reply_text(
            "❌ Invalid HEX color.\nExample: #1D6F42"
        )

        return

    result = format_excel(user_files[user], hex_color)

    await update.message.reply_document(document=open(result, "rb"))


# EXCEL FORMAT FUNCTION

def format_excel(file, header_color):

    df = pd.read_excel(file)

    output = "formatted.xlsx"

    df.to_excel(output, index=False)

    wb = load_workbook(output)
    ws = wb.active

    # add margin row and column
    ws.insert_rows(1)
    ws.insert_cols(1)

    thin = Side(style="thin", color="b7b7b7")

    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    header_fill = PatternFill(
        start_color=header_color.replace("#", ""),
        fill_type="solid"
    )

    gray_fill = PatternFill(start_color="F2F2F2", fill_type="solid")

    max_row = ws.max_row
    max_col = ws.max_column

    # HEADER ROW (row 2)

    for c in range(2, max_col + 1):

        cell = ws.cell(row=2, column=c)

        cell.fill = header_fill

        cell.font = Font(color="FFFFFF", bold=True)

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True
        )

        cell.border = border

    # DATA ROWS

    for r in range(3, max_row + 1):

        for c in range(2, max_col + 1):

            cell = ws.cell(row=r, column=c)

            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True
            )

            cell.border = border

            if r % 2 == 1:

                cell.fill = gray_fill

    # CLEAN MARGINS

    for c in range(1, max_col + 1):

        cell = ws.cell(row=1, column=c)

        cell.fill = PatternFill(fill_type=None)

        cell.border = Border()

    for r in range(1, max_row + 1):

        cell = ws.cell(row=r, column=1)

        cell.fill = PatternFill(fill_type=None)

        cell.border = Border()

    # keep left margin narrow

    ws.column_dimensions['A'].width = 2

    wb.save(output)

    return output


# BOT SETUP

app = ApplicationBuilder().token(TOKEN).build()

app.add_handler(CommandHandler("start", start))

app.add_handler(MessageHandler(filters.Document.ALL, handle_file))

app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, receive_hex))

app.add_handler(CallbackQueryHandler(default_color))

app.run_polling()
