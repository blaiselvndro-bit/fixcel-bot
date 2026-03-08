import os
import re
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    ApplicationBuilder,
    CommandHandler,
    MessageHandler,
    CallbackQueryHandler,
    filters,
    ContextTypes
)

TOKEN = os.getenv("BOT_TOKEN")

user_files = {}
user_colors = []
user_expected_colors = {}

# ---------- COLOR UTILS ----------

def hex_to_rgb(hex_color):
    hex_color = hex_color.replace("#", "")
    return tuple(int(hex_color[i:i+2], 16) for i in (0,2,4))

def rgb_to_hex(rgb):
    return "%02x%02x%02x" % rgb

def lighten(hex_color, factor=0.25):
    r,g,b = hex_to_rgb(hex_color)
    r = int(r + (255-r)*factor)
    g = int(g + (255-g)*factor)
    b = int(b + (255-b)*factor)
    return rgb_to_hex((r,g,b))

def darken(hex_color, factor=0.25):
    r,g,b = hex_to_rgb(hex_color)
    r = int(r*(1-factor))
    g = int(g*(1-factor))
    b = int(b*(1-factor))
    return rgb_to_hex((r,g,b))

def generate_palette(base_colors, count=12):
    palette = []
    for c in base_colors:
        palette.append(c)
        palette.append(lighten(c))
        palette.append(darken(c))
    while len(palette) < count:
        palette.append(lighten(palette[-1]))
    return palette

def is_dark(hex_color):
    r,g,b = hex_to_rgb(hex_color)
    brightness = (r*299 + g*587 + b*114)/1000
    return brightness < 140

# ---------- START ----------

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
"""
🎨 *Welcome to FIXCEL*

Send any Excel file and I will:

• Format tables
• Preserve charts
• Adapt Gantt charts
• Apply color themes
• Keep formulas intact
""",
parse_mode="Markdown"
)

# ---------- FILE UPLOAD ----------

async def handle_file(update: Update, context: ContextTypes.DEFAULT_TYPE):

    user = update.message.from_user.id
    file = update.message.document

    tg_file = await file.get_file()

    path = f"{user}_input.xlsx"

    await tg_file.download_to_drive(path)

    user_files[user] = path

    keyboard = [
        [InlineKeyboardButton("Excel Brand Color (#1D6F42)", callback_data="excel")],
        [InlineKeyboardButton("Custom Colors", callback_data="custom")]
    ]

    await update.message.reply_text(
        "Choose theme:",
        reply_markup=InlineKeyboardMarkup(keyboard)
    )

# ---------- COLOR STYLE ----------

async def choose_theme(update: Update, context: ContextTypes.DEFAULT_TYPE):

    query = update.callback_query
    await query.answer()

    user = query.from_user.id

    if query.data == "excel":

        result = format_excel(user_files[user], ["#1D6F42"])

        await query.message.reply_document(document=open(result,"rb"))
        return

    keyboard = [
        [InlineKeyboardButton("1",callback_data="c1"),
         InlineKeyboardButton("2",callback_data="c2"),
         InlineKeyboardButton("3",callback_data="c3")],
        [InlineKeyboardButton("4",callback_data="c4"),
         InlineKeyboardButton("5",callback_data="c5"),
         InlineKeyboardButton("6",callback_data="c6")]
    ]

    await query.message.reply_text(
        "How many colors?",
        reply_markup=InlineKeyboardMarkup(keyboard)
    )

# ---------- COLOR COUNT ----------

async def choose_color_count(update: Update, context: ContextTypes.DEFAULT_TYPE):

    query = update.callback_query
    await query.answer()

    user = query.from_user.id

    count = int(query.data[1])

    user_expected_colors[user] = count
    user_colors.clear()

    await query.message.reply_text("Send HEX color 1")

# ---------- RECEIVE HEX ----------

async def receive_hex(update: Update, context: ContextTypes.DEFAULT_TYPE):

    user = update.message.from_user.id

    if user not in user_files:
        return

    hex_color = update.message.text.strip()

    if not re.match(r'^#[0-9A-Fa-f]{6}$', hex_color):
        await update.message.reply_text("Invalid HEX example: #FF5733")
        return

    user_colors.append(hex_color)

    if len(user_colors) < user_expected_colors[user]:

        await update.message.reply_text(
            f"Send HEX color {len(user_colors)+1}"
        )
        return

    result = format_excel(user_files[user], user_colors)

    await update.message.reply_document(document=open(result,"rb"))

# ---------- CORE FORMATTER ----------

def format_excel(file, colors):

    wb = load_workbook(file)

    palette = generate_palette(colors)

    thin = Side(style="thin", color="b7b7b7")
    border = Border(left=thin,right=thin,top=thin,bottom=thin)

    for ws in wb.worksheets:

        max_row = ws.max_row
        max_col = ws.max_column

        font_color = "FFFFFF" if is_dark(colors[0]) else "333333"

        # Header styling
        for c in range(1,max_col+1):

            cell = ws.cell(row=1,column=c)

            if cell.value:

                color = palette[(c-1) % len(colors)]

                cell.fill = PatternFill(start_color=color.replace("#",""), fill_type="solid")

                cell.font = Font(color=font_color,bold=True)

                cell.alignment = Alignment(horizontal="center",vertical="center")

                cell.border = border

        # Data styling
        for r in range(2,max_row+1):

            for c in range(1,max_col+1):

                cell = ws.cell(row=r,column=c)

                if cell.value:

                    cell.alignment = Alignment(horizontal="center",vertical="center")

                    cell.border = border

        # Detect gantt bars (colored cells)
        for r in range(1,max_row+1):

            for c in range(1,max_col+1):

                cell = ws.cell(row=r,column=c)

                fill = cell.fill

                if fill and fill.start_color and fill.start_color.rgb:

                    new_color = palette[(r+c) % len(palette)]

                    cell.fill = PatternFill(start_color=new_color, fill_type="solid")

        # Move charts right if overlapping
        for chart in ws._charts:

            chart.anchor._from.col = max_col + 3
            chart.anchor._from.row = 2

    output = "formatted.xlsx"

    wb.save(output)

    return output

# ---------- BOT ----------

app = ApplicationBuilder().token(TOKEN).build()

app.add_handler(CommandHandler("start", start))
app.add_handler(MessageHandler(filters.Document.ALL, handle_file))
app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, receive_hex))

app.add_handler(CallbackQueryHandler(choose_theme, pattern="excel|custom"))
app.add_handler(CallbackQueryHandler(choose_color_count, pattern="c[1-6]"))

app.run_polling()
