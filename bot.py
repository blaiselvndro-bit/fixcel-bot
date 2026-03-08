import os
import re
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

user_files = {}
user_colors = {}


# ---------- COLOR HELPERS ----------

def hex_to_rgb(hex_color):
    hex_color = hex_color.replace("#", "")
    return tuple(int(hex_color[i:i+2], 16) for i in (0,2,4))


def rgb_to_hex(rgb):
    return "%02x%02x%02x" % rgb


def lighten_color(hex_color, factor=0.2):

    r,g,b = hex_to_rgb(hex_color)

    r = int(r + (255-r)*factor)
    g = int(g + (255-g)*factor)
    b = int(b + (255-b)*factor)

    return rgb_to_hex((r,g,b))


def darken_color(hex_color, factor=0.2):

    r,g,b = hex_to_rgb(hex_color)

    r = int(r*(1-factor))
    g = int(g*(1-factor))
    b = int(b*(1-factor))

    return rgb_to_hex((r,g,b))


def is_dark(hex_color):

    r,g,b = hex_to_rgb(hex_color)

    brightness = (r*299 + g*587 + b*114)/1000

    return brightness < 150


# ---------- START ----------

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):

    text = """
🎨 *Welcome to FIXCEL*

Send an Excel file and I will format it beautifully.

Features
• Smart table styling
• Alternating rows
• Chart color styling
• Wrap text
• Smart alignment

Free Plan
2 files per month

Premium
Unlimited formatting
$6/month
"""

    await update.message.reply_text(text, parse_mode="Markdown")


# ---------- FILE UPLOAD ----------

async def handle_file(update: Update, context: ContextTypes.DEFAULT_TYPE):

    user = update.message.from_user.id
    file = update.message.document

    tg_file = await file.get_file()

    path = f"{user}_input.xlsx"

    await tg_file.download_to_drive(path)

    user_files[user] = path

    await update.message.reply_text(
        "🎨 Send preferred HEX color for header.\nExample:\n#1D6F42"
    )


# ---------- RECEIVE HEX ----------

async def receive_hex(update: Update, context: ContextTypes.DEFAULT_TYPE):

    user = update.message.from_user.id

    if user not in user_files:
        return

    hex_color = update.message.text.strip()

    if not re.match(r'^#[0-9A-Fa-f]{6}$', hex_color):

        await update.message.reply_text("❌ Invalid HEX.\nExample: #1D6F42")
        return

    user_colors[user] = hex_color

    keyboard = [

        [
            InlineKeyboardButton("1 Color Theme", callback_data="chart1"),
            InlineKeyboardButton("2 Color Theme", callback_data="chart2")
        ]

    ]

    await update.message.reply_text(

        "📊 How many colors should charts use?",
        reply_markup=InlineKeyboardMarkup(keyboard)

    )


# ---------- CHART COLOR OPTION ----------

async def chart_choice(update: Update, context: ContextTypes.DEFAULT_TYPE):

    query = update.callback_query
    await query.answer()

    user = query.from_user.id

    color = user_colors[user]

    result = format_excel(user_files[user], color, query.data)

    await query.message.reply_document(document=open(result,"rb"))


# ---------- EXCEL FORMAT ----------

def format_excel(file, header_color, chart_mode):

    wb = load_workbook(file)

    ws = wb.active

    ws.insert_rows(1)
    ws.insert_cols(1)

    thin = Side(style="thin", color="b7b7b7")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    header_fill = PatternFill(start_color=header_color.replace("#",""), fill_type="solid")

    gray_fill = PatternFill(start_color="F2F2F2", fill_type="solid")
    white_fill = PatternFill(start_color="FFFFFF", fill_type="solid")

    max_row = ws.max_row
    max_col = ws.max_column

    header_font_color = "FFFFFF" if is_dark(header_color) else "333333"

    # HEADER

    for c in range(2, max_col+1):

        cell = ws.cell(row=2, column=c)

        cell.fill = header_fill
        cell.font = Font(color=header_font_color, bold=True)

        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        cell.border = border


    # DATA

    for r in range(3, max_row+1):

        for c in range(2, max_col+1):

            cell = ws.cell(row=r, column=c)

            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

            cell.border = border

            if r % 2 == 1:
                cell.fill = gray_fill
            else:
                cell.fill = white_fill


    # WHITE BACKGROUND

    for r in range(1, ws.max_row+50):

        for c in range(1, ws.max_column+50):

            if r >= 2 and c >= 2 and r <= max_row and c <= max_col:
                continue

            cell = ws.cell(row=r, column=c)

            cell.fill = white_fill
            cell.border = Border()


    ws.column_dimensions['A'].width = 2


    # ---------- CHART COLORING ----------

    if chart_mode == "chart1":

        base = header_color.replace("#","")

        shade1 = lighten_color(base,0.2)
        shade2 = darken_color(base,0.2)

    else:

        base = header_color.replace("#","")

        shade1 = lighten_color(base,0.3)
        shade2 = darken_color(base,0.3)


    for chart in ws._charts:

        try:

            for i,series in enumerate(chart.series):

                color = shade1 if i%2==0 else shade2

                series.graphicalProperties.solidFill = color

        except:
            pass


    output = "formatted.xlsx"

    wb.save(output)

    return output


# ---------- BOT SETUP ----------

app = ApplicationBuilder().token(TOKEN).build()

app.add_handler(CommandHandler("start", start))

app.add_handler(MessageHandler(filters.Document.ALL, handle_file))

app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, receive_hex))

app.add_handler(CallbackQueryHandler(chart_choice))

app.run_polling()
