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
user_color_count = {}

# ---------- COLOR HELPERS ----------

def hex_to_rgb(hex_color):
    hex_color = hex_color.replace("#","")
    return tuple(int(hex_color[i:i+2],16) for i in (0,2,4))

def rgb_to_hex(rgb):
    return "%02x%02x%02x" % rgb

def lighten(hex_color, factor=0.25):
    r,g,b = hex_to_rgb(hex_color)
    r=int(r+(255-r)*factor)
    g=int(g+(255-g)*factor)
    b=int(b+(255-b)*factor)
    return rgb_to_hex((r,g,b))

def darken(hex_color, factor=0.25):
    r,g,b = hex_to_rgb(hex_color)
    r=int(r*(1-factor))
    g=int(g*(1-factor))
    b=int(b*(1-factor))
    return rgb_to_hex((r,g,b))

def is_dark(hex_color):
    r,g,b = hex_to_rgb(hex_color)
    brightness=(r*299+g*587+b*114)/1000
    return brightness<140

# ---------- START ----------

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):

    await update.message.reply_text(
"""
🎨 *Welcome to FIXCEL*

Send an Excel file and I will format it beautifully.

✔ Tables supported
✔ Charts preserved
✔ Gantt charts supported
✔ Layout preserved
""",
parse_mode="Markdown"
)

# ---------- FILE UPLOAD ----------

async def handle_file(update: Update, context: ContextTypes.DEFAULT_TYPE):

    user=update.message.from_user.id
    file=update.message.document

    tg_file=await file.get_file()

    path=f"{user}_input.xlsx"

    await tg_file.download_to_drive(path)

    user_files[user]=path

    keyboard=[
        [InlineKeyboardButton("Use Excel Brand Color (#1D6F42)",callback_data="excel_color")],
        [InlineKeyboardButton("Use Custom Colors",callback_data="custom_colors")]
    ]

    await update.message.reply_text(
"🎨 Choose color style:",
reply_markup=InlineKeyboardMarkup(keyboard)
)

# ---------- COLOR STYLE ----------

async def color_choice(update: Update, context: ContextTypes.DEFAULT_TYPE):

    query=update.callback_query
    await query.answer()

    user=query.from_user.id

    if query.data=="excel_color":

        result=format_excel(user_files[user],["#1D6F42"])

        await query.message.reply_document(document=open(result,"rb"))
        return

    keyboard=[
        [InlineKeyboardButton("1",callback_data="c1"),InlineKeyboardButton("2",callback_data="c2"),InlineKeyboardButton("3",callback_data="c3")],
        [InlineKeyboardButton("4",callback_data="c4"),InlineKeyboardButton("5",callback_data="c5"),InlineKeyboardButton("6",callback_data="c6")]
    ]

    await query.message.reply_text(
"How many header colors should be used?",
reply_markup=InlineKeyboardMarkup(keyboard)
)

# ---------- COLOR COUNT ----------

async def color_count(update: Update, context: ContextTypes.DEFAULT_TYPE):

    query=update.callback_query
    await query.answer()

    user=query.from_user.id

    count=int(query.data[1])

    user_color_count[user]=count
    user_colors[user]=[]

    await query.message.reply_text(f"Send HEX color 1 of {count}")

# ---------- RECEIVE HEX COLORS ----------

async def receive_hex(update: Update, context: ContextTypes.DEFAULT_TYPE):

    user=update.message.from_user.id

    if user not in user_files:
        return

    hex_color=update.message.text.strip()

    if not re.match(r'^#[0-9A-Fa-f]{6}$', hex_color):

        await update.message.reply_text("Invalid HEX example: #FF5733")
        return

    user_colors[user].append(hex_color)

    if len(user_colors[user])<user_color_count[user]:

        await update.message.reply_text(
f"Send HEX color {len(user_colors[user])+1} of {user_color_count[user]}"
)
        return

    result=format_excel(user_files[user],user_colors[user])

    await update.message.reply_document(document=open(result,"rb"))

# ---------- FORMAT EXCEL ----------

def format_excel(file,colors):

    wb=load_workbook(file)

    thin=Side(style="thin",color="b7b7b7")
    border=Border(left=thin,right=thin,top=thin,bottom=thin)

    gray_fill=PatternFill(start_color="F2F2F2",fill_type="solid")

    for ws in wb.worksheets:

        max_row=ws.max_row
        max_col=ws.max_column

        font_color="FFFFFF" if is_dark(colors[0]) else "333333"

        # HEADER STYLE
        for c in range(1,max_col+1):

            cell=ws.cell(row=1,column=c)

            if cell.value:

                color=colors[(c-1)%len(colors)].replace("#","")

                cell.fill=PatternFill(start_color=color,fill_type="solid")

                cell.font=Font(color=font_color,bold=True)

                cell.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)

                cell.border=border

        # DATA STYLE
        for r in range(2,max_row+1):

            for c in range(1,max_col+1):

                cell=ws.cell(row=r,column=c)

                if cell.value:

                    cell.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)

                    cell.border=border

                    if r%2==0:
                        cell.fill=gray_fill

        # GANTT BAR ADAPTATION
        for r in range(1,max_row+1):

            for c in range(1,max_col+1):

                cell=ws.cell(row=r,column=c)

                if cell.fill and cell.fill.start_color and cell.fill.start_color.rgb:

                    base=colors[(r+c)%len(colors)]

                    new_color=lighten(base,0.3)

                    cell.fill=PatternFill(start_color=new_color,fill_type="solid")

        # CHART POSITIONING
        for chart in ws._charts:

            chart.anchor._from.col=max_col+2
            chart.anchor._from.row=1

    output="formatted.xlsx"

    wb.save(output)

    return output

# ---------- BOT ----------

app=ApplicationBuilder().token(TOKEN).build()

app.add_handler(CommandHandler("start",start))
app.add_handler(MessageHandler(filters.Document.ALL,handle_file))
app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND,receive_hex))
app.add_handler(CallbackQueryHandler(color_choice,pattern="excel_color|custom_colors"))
app.add_handler(CallbackQueryHandler(color_count,pattern="c[1-6]"))

app.run_polling()
