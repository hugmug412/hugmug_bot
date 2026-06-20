#!/usr/bin/env python3
"""
ربات حسابداری فروش تلگرام
Personal Sales Accounting Telegram Bot
"""

import os
import json
import sqlite3
from datetime import datetime, timedelta
from telegram import Update, ReplyKeyboardMarkup, KeyboardButton
from telegram.ext import (
    Application, CommandHandler, MessageHandler,
    filters, ContextTypes, ConversationHandler
)

# ===== تنظیمات =====
BOT_TOKEN = os.environ.get("BOT_TOKEN")  # توکن از متغیر محیطی خونده می‌شه (امن‌تر)
if not BOT_TOKEN:
    raise ValueError("متغیر محیطی BOT_TOKEN تنظیم نشده! آن را در تنظیمات سرور اضافه کن.")
DB_FILE = "sales.db"

# States for conversation
WAITING_PRODUCT = 1
WAITING_AMOUNT = 2
WAITING_QUANTITY = 3

# ===== دیتابیس =====
def init_db():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute('''
        CREATE TABLE IF NOT EXISTS sales (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            product TEXT NOT NULL,
            amount REAL NOT NULL,
            quantity INTEGER DEFAULT 1,
            date TEXT NOT NULL,
            user_id INTEGER NOT NULL
        )
    ''')
    conn.commit()
    conn.close()

def add_sale(product: str, amount: float, quantity: int, user_id: int):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute(
        "INSERT INTO sales (product, amount, quantity, date, user_id) VALUES (?, ?, ?, ?, ?)",
        (product, amount, quantity, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), user_id)
    )
    conn.commit()
    conn.close()

def get_sales_in_range(user_id: int, start_date: str, end_date: str):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute(
        "SELECT product, amount, quantity, date FROM sales WHERE user_id=? AND date BETWEEN ? AND ? ORDER BY date DESC",
        (user_id, start_date, end_date)
    )
    rows = c.fetchall()
    conn.close()
    return rows

def get_product_summary(user_id: int, start_date: str, end_date: str):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute(
        """SELECT product, SUM(amount*quantity) as total, SUM(quantity) as qty
           FROM sales WHERE user_id=? AND date BETWEEN ? AND ?
           GROUP BY product ORDER BY total DESC""",
        (user_id, start_date, end_date)
    )
    rows = c.fetchall()
    conn.close()
    return rows

# ===== کیبورد =====
def main_keyboard():
    buttons = [
        [KeyboardButton("➕ ثبت فروش"), KeyboardButton("📊 گزارش هفتگی")],
        [KeyboardButton("📈 گزارش ماهانه"), KeyboardButton("🏆 برترین محصولات")],
        [KeyboardButton("📋 آخرین فروش‌ها"), KeyboardButton("💰 امروز")],
    ]
    return ReplyKeyboardMarkup(buttons, resize_keyboard=True)

# ===== هندلرها =====
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    name = update.effective_user.first_name
    await update.message.reply_text(
        f"سلام {name}! 👋\n\n"
        "به ربات حسابداری فروشت خوش اومدی 📦💰\n\n"
        "از منوی پایین انتخاب کن:",
        reply_markup=main_keyboard()
    )

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text
    
    if text == "➕ ثبت فروش":
        await update.message.reply_text("📦 نام محصول فروخته‌شده رو بنویس:")
        return WAITING_PRODUCT
    
    elif text == "📊 گزارش هفتگی":
        await send_weekly_report(update, context)
    
    elif text == "📈 گزارش ماهانه":
        await send_monthly_report(update, context)
    
    elif text == "🏆 برترین محصولات":
        await send_top_products(update, context)
    
    elif text == "📋 آخرین فروش‌ها":
        await send_recent_sales(update, context)
    
    elif text == "💰 امروز":
        await send_today_report(update, context)
    
    return ConversationHandler.END

async def get_product(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data['product'] = update.message.text
    await update.message.reply_text(
        f"✅ محصول: {update.message.text}\n\n"
        "💵 مبلغ فروش رو وارد کن (به تومان):"
    )
    return WAITING_AMOUNT

async def get_amount(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        amount_text = update.message.text.replace(",", "").replace("،", "")
        amount = float(amount_text)
        context.user_data['amount'] = amount
        await update.message.reply_text(
            f"✅ مبلغ: {amount:,.0f} تومان\n\n"
            "🔢 تعداد فروخته‌شده رو بنویس (اگه ۱ تاست، عدد ۱ رو بزن):"
        )
        return WAITING_QUANTITY
    except ValueError:
        await update.message.reply_text("❌ لطفاً فقط عدد وارد کن (مثل: 150000)")
        return WAITING_AMOUNT

async def get_quantity(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        qty = int(update.message.text)
        if qty <= 0:
            raise ValueError
        
        product = context.user_data['product']
        amount = context.user_data['amount']
        user_id = update.effective_user.id
        
        add_sale(product, amount, qty, user_id)
        
        total = amount * qty
        await update.message.reply_text(
            f"✅ فروش ثبت شد!\n\n"
            f"📦 محصول: {product}\n"
            f"💵 قیمت واحد: {amount:,.0f} تومان\n"
            f"🔢 تعداد: {qty}\n"
            f"💰 جمع کل: {total:,.0f} تومان\n"
            f"🕐 زمان: {datetime.now().strftime('%H:%M - %Y/%m/%d')}",
            reply_markup=main_keyboard()
        )
        context.user_data.clear()
        return ConversationHandler.END
        
    except ValueError:
        await update.message.reply_text("❌ لطفاً یه عدد درست وارد کن:")
        return WAITING_QUANTITY

async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    await update.message.reply_text("❌ لغو شد.", reply_markup=main_keyboard())
    return ConversationHandler.END

# ===== گزارش‌ها =====
def format_number(n):
    return f"{n:,.0f}"

async def send_today_report(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    today = datetime.now().strftime("%Y-%m-%d")
    start = today + " 00:00:00"
    end = today + " 23:59:59"
    
    sales = get_sales_in_range(user_id, start, end)
    
    if not sales:
        await update.message.reply_text("📭 امروز هیچ فروشی ثبت نشده.")
        return
    
    total = sum(s[1] * s[2] for s in sales)
    count = sum(s[2] for s in sales)
    
    text = f"☀️ گزارش امروز - {today}\n"
    text += "━━━━━━━━━━━━━━━━\n"
    
    for p, a, q, d in sales:
        time = d.split(" ")[1][:5]
        text += f"🕐 {time} | {p} × {q} = {format_number(a*q)} ت\n"
    
    text += "━━━━━━━━━━━━━━━━\n"
    text += f"🛒 تعداد آیتم: {count}\n"
    text += f"💰 جمع کل: {format_number(total)} تومان"
    
    await update.message.reply_text(text)

async def send_weekly_report(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    end = datetime.now()
    start = end - timedelta(days=7)
    
    sales = get_sales_in_range(user_id,
        start.strftime("%Y-%m-%d 00:00:00"),
        end.strftime("%Y-%m-%d 23:59:59")
    )
    products = get_product_summary(user_id,
        start.strftime("%Y-%m-%d 00:00:00"),
        end.strftime("%Y-%m-%d 23:59:59")
    )
    
    if not sales:
        await update.message.reply_text("📭 در هفته گذشته هیچ فروشی ثبت نشده.")
        return
    
    total = sum(s[1] * s[2] for s in sales)
    total_items = sum(s[2] for s in sales)
    
    text = f"📊 گزارش هفتگی\n"
    text += f"📅 {start.strftime('%Y/%m/%d')} تا {end.strftime('%Y/%m/%d')}\n"
    text += "━━━━━━━━━━━━━━━━\n"
    text += f"🛒 کل فروش: {len(sales)} سفارش ({total_items} آیتم)\n"
    text += f"💰 درآمد کل: {format_number(total)} تومان\n"
    text += f"📌 میانگین روزانه: {format_number(total/7)} تومان\n"
    text += "\n🏆 فروش محصولات:\n"
    
    for i, (prod, tot, qty) in enumerate(products[:5], 1):
        text += f"{i}. {prod}: {qty} عدد = {format_number(tot)} ت\n"
    
    await update.message.reply_text(text)

async def send_monthly_report(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    now = datetime.now()
    start = now.replace(day=1, hour=0, minute=0, second=0)
    
    sales = get_sales_in_range(user_id,
        start.strftime("%Y-%m-%d 00:00:00"),
        now.strftime("%Y-%m-%d 23:59:59")
    )
    products = get_product_summary(user_id,
        start.strftime("%Y-%m-%d 00:00:00"),
        now.strftime("%Y-%m-%d 23:59:59")
    )
    
    if not sales:
        await update.message.reply_text("📭 این ماه هیچ فروشی ثبت نشده.")
        return
    
    total = sum(s[1] * s[2] for s in sales)
    total_items = sum(s[2] for s in sales)
    days_passed = (now - start).days + 1
    
    text = f"📈 گزارش ماهانه\n"
    text += f"📅 {start.strftime('%Y/%m')} | {days_passed} روز گذشته\n"
    text += "━━━━━━━━━━━━━━━━\n"
    text += f"🛒 کل سفارشات: {len(sales)} ({total_items} آیتم)\n"
    text += f"💰 درآمد کل: {format_number(total)} تومان\n"
    text += f"📌 میانگین روزانه: {format_number(total/days_passed)} تومان\n"
    text += f"📆 پیش‌بینی ماهانه: {format_number(total/days_passed*30)} تومان\n"
    text += "\n🏆 برترین محصولات ماه:\n"
    
    for i, (prod, tot, qty) in enumerate(products, 1):
        pct = (tot / total * 100) if total else 0
        text += f"{i}. {prod}\n   {qty} عدد | {format_number(tot)} ت ({pct:.0f}%)\n"
    
    await update.message.reply_text(text)

async def send_top_products(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    end = datetime.now()
    start = end - timedelta(days=30)
    
    products = get_product_summary(user_id,
        start.strftime("%Y-%m-%d 00:00:00"),
        end.strftime("%Y-%m-%d 23:59:59")
    )
    
    if not products:
        await update.message.reply_text("📭 داده‌ای برای نمایش وجود نداره.")
        return
    
    total_all = sum(p[1] for p in products)
    
    medals = ["🥇", "🥈", "🥉", "4️⃣", "5️⃣"]
    text = "🏆 برترین محصولات (۳۰ روز اخیر)\n"
    text += "━━━━━━━━━━━━━━━━\n"
    
    for i, (prod, tot, qty) in enumerate(products):
        medal = medals[i] if i < len(medals) else f"{i+1}."
        pct = (tot / total_all * 100) if total_all else 0
        bar = "█" * int(pct / 10) + "░" * (10 - int(pct / 10))
        text += f"{medal} {prod}\n"
        text += f"   {bar} {pct:.0f}%\n"
        text += f"   {qty} عدد | {format_number(tot)} تومان\n\n"
    
    await update.message.reply_text(text)

async def send_recent_sales(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    end = datetime.now()
    start = end - timedelta(days=30)
    
    sales = get_sales_in_range(user_id,
        start.strftime("%Y-%m-%d 00:00:00"),
        end.strftime("%Y-%m-%d 23:59:59")
    )[:10]
    
    if not sales:
        await update.message.reply_text("📭 فروشی ثبت نشده.")
        return
    
    text = "📋 آخرین ۱۰ فروش\n"
    text += "━━━━━━━━━━━━━━━━\n"
    
    for prod, amt, qty, date in sales:
        d = date.split(" ")[0]
        t = date.split(" ")[1][:5]
        text += f"📦 {prod} × {qty}\n"
        text += f"   💵 {format_number(amt*qty)} ت | 📅 {d} {t}\n\n"
    
    await update.message.reply_text(text)

# ===== اجرا =====
def main():
    init_db()
    app = Application.builder().token(BOT_TOKEN).build()
    
    conv_handler = ConversationHandler(
        entry_points=[MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message)],
        states={
            WAITING_PRODUCT: [MessageHandler(filters.TEXT & ~filters.COMMAND, get_product)],
            WAITING_AMOUNT: [MessageHandler(filters.TEXT & ~filters.COMMAND, get_amount)],
            WAITING_QUANTITY: [MessageHandler(filters.TEXT & ~filters.COMMAND, get_quantity)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
    )
    
    app.add_handler(CommandHandler("start", start))
    app.add_handler(conv_handler)
    
    print("✅ ربات حسابداری شروع به کار کرد...")
    app.run_polling(allowed_updates=Update.ALL_TYPES)

if __name__ == "__main__":
    main()
