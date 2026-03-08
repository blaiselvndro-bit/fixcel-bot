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
from openpyxl.utils import column_index_from_string, get_column_letter

TOKEN = os.getenv("BOT_TOKEN")

user_files = {}
user_colors = {}
user_color_count = {}

# ---------- COLOR HELPERS ----------

def hex_to_rgb(hex_color):
    hex_color = hex_color.replace("#","")
    return tuple(int(hex_color[i:i+2],16) for i in (0,2,4))

def is_dark(hex_color):
    r,g,b = hex_to_rgb(hex_color)
    brightness=(r*299+g*587+b*114)/1000
    return brightness<140

# ---------- CHART REFERENCE SHIFT ----------

def shift_reference(ref):

    pattern=r'(\$?[A-Z]+)(\$?\d+)'

    def repl(match):

        col=match.group(1).replace("$","")
        row=int(match.group(2).replace("$",""))

        new_col=get_column_letter(column_index_from_string(col)+1)
        new_row=row+1

        return f"${new_col}${new_row}"

    return re.sub(pattern,repl,ref)

# ---------- START ----------

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):

    await update.message.reply_text(
"""
🎨 *Welcome to FIXCEL*

Send an Excel file and I will format it beautifully.

Charts, legends, and formulas will be preserved.
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

    if not hex_color.startswith("#") or len(hex_color)!=7:

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
    ws=wb.active

    # shift chart references FIRST
    for chart in ws._charts:

        try:
            for series in chart.series:

                if series.val and series.val.numRef:

                    series.val.numRef.f=shift_reference(series.val.numRef.f)

                if series.cat and series.cat.numRef:

                    series.cat.numRef.f=shift_reference(series.cat.numRef.f)

        except:
            pass

    # insert margins
    ws.insert_rows(1)
    ws.insert_cols(1)

    thin=Side(style="thin",color="b7b7b7")
    border=Border(left=thin,right=thin,top=thin,bottom=thin)

    gray_fill=PatternFill(start_color="F2F2F2",fill_type="solid")
    white_fill=PatternFill(start_color="FFFFFF",fill_type="solid")

    max_row=ws.max_row
    max_col=ws.max_column

    font_color="FFFFFF" if is_dark(colors[0]) else "333333"

    # HEADER

    for c in range(2,max_col+1):

        color=colors[(c-2)%len(colors)].replace("#","")

        cell=ws.cell(row=2,column=c)

        cell.fill=PatternFill(start_color=color,fill_type="solid")
        cell.font=Font(color=font_color,bold=True)
        cell.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
        cell.border=border

    # DATA

    for r in range(3,max_row+1):

        for c in range(2,max_col+1):

            cell=ws.cell(row=r,column=c)

            cell.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
            cell.border=border

            if r%2==1:
                cell.fill=gray_fill
            else:
                cell.fill=white_fill

    # CLEAN OUTSIDE

    for r in range(1,max_row+30):

        for c in range(1,max_col+30):

            if r>=2 and c>=2 and r<=max_row and c<=max_col:
                continue

            cell=ws.cell(row=r,column=c)
            cell.fill=white_fill
            cell.border=Border()

    # move charts right
    for chart in ws._charts:

        chart.anchor._from.col=max_col+3
        chart.anchor._from.row=2

    ws.column_dimensions['A'].width=2

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
