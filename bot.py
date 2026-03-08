import os
import pandas as pd
import sqlite3
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, filters, ContextTypes
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment

TOKEN = os.getenv("BOT_TOKEN")

# ---------- DATABASE ----------

def init_db():
    conn = sqlite3.connect("users.db")
    cur = conn.cursor()

    cur.execute("""
    CREATE TABLE IF NOT EXISTS users(
        user_id INTEGER PRIMARY KEY,
        files_used INTEGER DEFAULT 0,
        month TEXT,
        premium INTEGER DEFAULT 0
    )
    """)

    conn.commit()
    conn.close()

init_db()


# ---------- FORMAT FUNCTION ----------

def format_excel(file, header_color):

    df = pd.read_excel(file)

    output = "formatted.xlsx"
    df.to_excel(output, index=False)

    wb = load_workbook(output)
    ws = wb.active

    header_fill = PatternFill(start_color=header_color.replace("#",""), fill_type="solid")

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center", wrap_text=True)

    gray_fill = PatternFill(start_color="F2F2F2", fill_type="solid")

    for i,row in enumerate(ws.iter_rows(min_row=2), start=2):

        for cell in row:

            cell.alignment = Alignment(horizontal="left", wrap_text=True)

            if i % 2 == 0:
                cell.fill = gray_fill

    wb.save(output)

    return output


# ---------- USER LIMIT CHECK ----------

def can_use(user_id):

    conn = sqlite3.connect("users.db")
    cur = conn.cursor()

    cur.execute("SELECT files_used,premium FROM users WHERE user_id=?", (user_id,))
    row = cur.fetchone()

    if row is None:
        cur.execute("INSERT INTO users(user_id,files_used,premium) VALUES (?,0,0)",(user_id,))
        conn.commit()
        conn.close()
        return True

    files_used, premium = row

    if premium == 1:
        conn.close()
        return True

    if files_used >= 2:
        conn.close()
        return False

    cur.execute("UPDATE users SET files_used=files_used+1 WHERE user_id=?", (user_id,))
    conn.commit()
    conn.close()

    return True


# ---------- BOT COMMANDS ----------

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):

    await update.message.reply_text(
"""
Welcome to FIXCEL 🎨

Upload an Excel file and I will format it beautifully.

Features:
• Colored header
• Alternating rows
• Wrap text
• Clean alignment

Free Plan
2 files per month

Premium
Unlimited formatting
$6/month
"""
)


async def handle_file(update: Update, context: ContextTypes.DEFAULT_TYPE):

    user = update.message.from_user.id

    if not can_use(user):

        await update.message.reply_text(
"⚠️ Free limit reached.\nUpgrade to premium for unlimited formatting.\n\nUse /premium"
        )
        return

    file = update.message.document

    tg_file = await file.get_file()

    await tg_file.download_to_drive("input.xlsx")

    header_color = "#1D6F42"

    result = format_excel("input.xlsx", header_color)

    await update.message.reply_document(document=open(result,"rb"))


# ---------- PREMIUM PAYMENT ----------

async def premium(update: Update, context: ContextTypes.DEFAULT_TYPE):

    await update.message.reply_text(
"""
⭐ FIXCEL Premium

Unlimited Excel formatting.

Price:
$6 per month via Telegram Stars.
"""
)


# ---------- START BOT ----------

app = ApplicationBuilder().token(TOKEN).build()

app.add_handler(CommandHandler("start", start))
app.add_handler(CommandHandler("premium", premium))
app.add_handler(MessageHandler(filters.Document.ALL, handle_file))

app.run_polling()
