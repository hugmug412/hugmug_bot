#!/usr/bin/env python3
"""
ربات حسابداری فروش تلگرام - نسخه ۲
Personal Sales Accounting Telegram Bot - v2
"""

import os
import sqlite3
from datetime import datetime, timedelta
from telegram import (
    Update, ReplyKeyboardMarkup, KeyboardButton,
    InlineKeyboardMarkup, InlineKeyboardButton
)
from telegram.ext import (
    Application, CommandHandler, MessageHandler,
    filters, ContextTypes, ConversationHandler, CallbackQueryHandler
)
from reportlab.lib.pagesizes import A5
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.enums import TA_RIGHT
from reportlab.lib.utils import ImageReader
import arabic_reshaper
from bidi.algorithm import get_display
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment

# ===== تنظیمات =====
BOT_TOKEN = os.environ.get("BOT_TOKEN")
if not BOT_TOKEN:
    raise ValueError("متغیر محیطی BOT_TOKEN تنظیم نشده! آن را در تنظیمات سرور اضافه کن.")
DB_FILE = "sales.db"
PROFIT_MARGIN = 0.30  # درصد سود تقریبی (۳۰٪)

# اطلاعات ثابت فروشگاه (فرستنده)
STORE_NAME = "فروشگاه لوازم کادویی هاگ ماگ"
STORE_ADDRESS = "خوزستان، اندیمشک، خیابان سینا، جنب آزمایشگاه سینا"
STORE_PHONE = "09376923487"
LOGO_PATH = "logo.png"

# States for conversation
WAITING_QUICK_SALE = 1
WAITING_LABEL_INFO = 2
WAITING_ORDER_ADDRESS = 3

# ===== فونت فارسی برای PDF =====
FONT_NAME = "Vazir"
FONT_REGISTERED = False

def register_persian_font():
    global FONT_REGISTERED
    if FONT_REGISTERED:
        return
    # دانلود فونت فارسی در صورت نبود
    font_path = "Vazirmatn-Regular.ttf"
    if not os.path.exists(font_path):
        import urllib.request
        url = "https://github.com/rastikerdar/vazirmatn/raw/master/fonts/ttf/Vazirmatn-Regular.ttf"
        try:
            urllib.request.urlretrieve(url, font_path)
        except Exception:
            pass
    if os.path.exists(font_path):
        pdfmetrics.registerFont(TTFont(FONT_NAME, font_path))
        FONT_REGISTERED = True

def fa(text):
    """تبدیل متن فارسی برای نمایش درست در PDF (راست‌چین و چسبیده)"""
    reshaped = arabic_reshaper.reshape(text)
    return get_display(reshaped)

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
    sale_id = c.lastrowid
    conn.close()
    return sale_id

def delete_sale(sale_id: int, user_id: int):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("DELETE FROM sales WHERE id=? AND user_id=?", (sale_id, user_id))
    conn.commit()
    deleted = c.rowcount > 0
    conn.close()
    return deleted

def get_sale_by_id(sale_id: int, user_id: int):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("SELECT id, product, amount, quantity, date FROM sales WHERE id=? AND user_id=?", (sale_id, user_id))
    row = c.fetchone()
    conn.close()
    return row

def update_sale(sale_id: int, user_id: int, field: str, value):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute(f"UPDATE sales SET {field}=? WHERE id=? AND user_id=?", (value, sale_id, user_id))
    conn.commit()
    updated = c.rowcount > 0
    conn.close()
    return updated

def get_sales_in_range(user_id: int, start_date: str, end_date: str):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute(
        "SELECT id, product, amount, quantity, date FROM sales WHERE user_id=? AND date BETWEEN ? AND ? ORDER BY date DESC",
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
        [KeyboardButton("➕ ثبت فروش"), KeyboardButton("✏️ ویرایش/حذف فروش")],
        [KeyboardButton("💰 امروز"), KeyboardButton("📊 گزارش هفتگی")],
        [KeyboardButton("📈 گزارش ماهانه"), KeyboardButton("🏆 برترین محصولات")],
        [KeyboardButton("📋 آخرین فروش‌ها"), KeyboardButton("🏷 ساخت فاکتور پستی")],
        [KeyboardButton("📥 خروجی اکسل ماه")],
    ]
    return ReplyKeyboardMarkup(buttons, resize_keyboard=True)

# ===== هندلر شروع =====
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    name = update.effective_user.first_name
    await update.message.reply_text(
        f"سلام {name}! 👋\n\n"
        "به ربات حسابداری فروشت خوش اومدی 📦💰\n\n"
        "⚡️ ثبت سریع: کافیه این‌جوری بنویسی:\n"
        "نام محصول، تعداد، قیمت واحد\n"
        "مثال: دستمال مرطوب، ۲، ۱۵۰۰۰۰\n\n"
        "یا از منوی پایین استفاده کن:",
        reply_markup=main_keyboard()
    )

# ===== ثبت سریع فروش با پیام تکی =====
def try_parse_quick_sale(text: str):
    """تلاش برای پارس کردن پیام به فرم: نام، تعداد، قیمت"""
    parts = [p.strip() for p in text.replace("،", ",").split(",")]
    if len(parts) != 3:
        return None
    product, qty_str, amount_str = parts
    try:
        qty = int(qty_str.replace(" ", ""))
        amount = float(amount_str.replace(" ", "").replace(",", ""))
        if qty <= 0 or amount < 0 or not product:
            return None
        return product, qty, amount
    except ValueError:
        return None

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text
    user_id = update.effective_user.id

    # اگه منتظر مقدار جدید برای ویرایش هستیم
    if context.user_data.get('awaiting_edit_value'):
        await process_edit_value(update, context)
        return ConversationHandler.END

    # اگه منتظر آدرس سفارش (فلوی سریع فروش+آدرس) هستیم
    if context.user_data.get('awaiting_order_address'):
        await process_order_address(update, context)
        return ConversationHandler.END

    if text == "➕ ثبت فروش":
        await update.message.reply_text(
            "📦 فروش رو این‌جوری بفرست:\n"
            "نام محصول، تعداد، قیمت واحد\n\n"
            "مثال:\nدستمال مرطوب، ۲، ۱۵۰۰۰۰\n\n"
            "(برای لغو /cancel رو بزن)"
        )
        return WAITING_QUICK_SALE

    elif text == "✏️ ویرایش/حذف فروش":
        await send_delete_menu(update, context)

    elif text == "📥 خروجی اکسل ماه":
        await send_excel_export(update, context)

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

    elif text == "🏷 ساخت فاکتور پستی":
        await update.message.reply_text(
            "🏷 مشخصات گیرنده رو بفرست (هر چی هست، همونجوری پیست کن - نام، آدرس، تلفن، کدپستی):\n\n"
            "بعدش محصولات و روش ارسال رو می‌پرسم.\n\n"
            "(برای لغو /cancel رو بزن)"
        )
        context.user_data['standalone_label'] = True
        return WAITING_LABEL_INFO

    else:
        # تلاش برای ثبت سریع فروش بدون نیاز به دکمه
        parsed = try_parse_quick_sale(text)
        if parsed:
            product, qty, amount = parsed
            sale_id = add_sale(product, amount, qty, user_id)
            total = amount * qty
            context.user_data['last_order'] = {
                'sale_id': sale_id, 'product': product, 'qty': qty, 'amount': amount
            }
            keyboard = [[
                InlineKeyboardButton("🏷 بله، آدرس رو می‌فرستم", callback_data="addaddr_yes"),
                InlineKeyboardButton("❌ نه، فقط ثبت فروش", callback_data="addaddr_no"),
            ]]
            await update.message.reply_text(
                f"✅ فروش ثبت شد!\n\n"
                f"📦 محصول: {product}\n"
                f"💵 قیمت واحد: {amount:,.0f} تومان\n"
                f"🔢 تعداد: {qty}\n"
                f"💰 جمع کل: {total:,.0f} تومان\n"
                f"🕐 زمان: {datetime.now().strftime('%H:%M - %Y/%m/%d')}\n\n"
                f"📮 می‌خوای فاکتور پستی هم براش بسازم؟",
                reply_markup=InlineKeyboardMarkup(keyboard)
            )

    return ConversationHandler.END

async def get_quick_sale(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text
    user_id = update.effective_user.id
    parsed = try_parse_quick_sale(text)

    if not parsed:
        await update.message.reply_text(
            "❌ فرمت درست نیست. دوباره امتحان کن:\n"
            "نام محصول، تعداد، قیمت واحد\n"
            "مثال: دستمال مرطوب، ۲، ۱۵۰۰۰۰"
        )
        return WAITING_QUICK_SALE

    product, qty, amount = parsed
    sale_id = add_sale(product, amount, qty, user_id)
    total = amount * qty
    context.user_data['last_order'] = {
        'sale_id': sale_id, 'product': product, 'qty': qty, 'amount': amount
    }
    keyboard = [[
        InlineKeyboardButton("🏷 بله، آدرس رو می‌فرستم", callback_data="addaddr_yes"),
        InlineKeyboardButton("❌ نه، فقط ثبت فروش", callback_data="addaddr_no"),
    ]]
    await update.message.reply_text(
        f"✅ فروش ثبت شد!\n\n"
        f"📦 محصول: {product}\n"
        f"💵 قیمت واحد: {amount:,.0f} تومان\n"
        f"🔢 تعداد: {qty}\n"
        f"💰 جمع کل: {total:,.0f} تومان\n"
        f"🕐 زمان: {datetime.now().strftime('%H:%M - %Y/%m/%d')}\n\n"
        f"📮 می‌خوای فاکتور پستی هم براش بسازم؟",
        reply_markup=InlineKeyboardMarkup(keyboard)
    )
    return ConversationHandler.END

async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    await update.message.reply_text("❌ لغو شد.", reply_markup=main_keyboard())
    return ConversationHandler.END

# ===== ویرایش / حذف فروش =====
async def send_delete_menu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    end = datetime.now()
    start = end - timedelta(days=30)

    sales = get_sales_in_range(user_id,
        start.strftime("%Y-%m-%d 00:00:00"),
        end.strftime("%Y-%m-%d 23:59:59")
    )[:15]

    if not sales:
        await update.message.reply_text("📭 فروشی برای ویرایش یا حذف وجود نداره.")
        return

    for sale_id, prod, amt, qty, date in sales:
        d = date.split(" ")[0][5:]  # ماه-روز
        keyboard = [[
            InlineKeyboardButton("✏️ ویرایش", callback_data=f"edit_{sale_id}"),
            InlineKeyboardButton("🗑 حذف", callback_data=f"del_{sale_id}")
        ]]
        await update.message.reply_text(
            f"📦 {prod} × {qty} | 💰 {format_number(amt*qty)}ت | 📅 {d}",
            reply_markup=InlineKeyboardMarkup(keyboard)
        )

async def handle_delete_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    user_id = update.effective_user.id

    if query.data == "addaddr_yes":
        await query.edit_message_text(
            "📮 مشخصات گیرنده رو بفرست (هر چی هست، همونجوری پیست کن):\n\n"
            "(برای لغو /cancel رو بزن)"
        )
        context.user_data['awaiting_order_address'] = True
        return

    elif query.data == "addaddr_no":
        await query.edit_message_text("✅ باشه، فقط فروش ثبت موند.")
        context.user_data.pop('last_order', None)
        return

    elif query.data.startswith("shipvia_"):
        method = query.data.split("_")[1]
        method_names = {"tipax": "تیپاکس", "post": "پست", "chapar": "چاپار"}
        method_fa = method_names.get(method, method)
        order = context.user_data.get('pending_label_order')

        if not order:
            await query.edit_message_text("❌ اطلاعات سفارش پیدا نشد، دوباره امتحان کن.")
            return

        try:
            pdf_path = create_shipping_label(
                order['recipient_info'], order['products_line'], method_fa
            )
            with open(pdf_path, "rb") as f:
                await context.bot.send_document(
                    chat_id=update.effective_chat.id,
                    document=f,
                    filename="فاکتور_پستی.pdf",
                    caption=f"🏷 فاکتور پستی ({method_fa}) ساخته شد ✅"
                )
            await query.edit_message_text(f"✅ روش ارسال: {method_fa} - فاکتور ارسال شد.")
        except Exception as e:
            await query.edit_message_text(f"❌ خطا در ساخت فاکتور: {e}")

        context.user_data.pop('pending_label_order', None)
        return

    if query.data.startswith("del_"):
        sale_id = int(query.data.split("_")[1])
        sale = get_sale_by_id(sale_id, user_id)
        if not sale:
            await query.edit_message_text("❌ این فروش پیدا نشد (شاید قبلاً حذف شده).")
            return
        keyboard = [[
            InlineKeyboardButton("✅ بله، حذف کن", callback_data=f"confirmdel_{sale_id}"),
            InlineKeyboardButton("❌ نه، بیخیال", callback_data="canceldel")
        ]]
        _, prod, amt, qty, date = sale
        await query.edit_message_text(
            f"⚠️ مطمئنی می‌خوای این فروش رو حذف کنی؟\n\n"
            f"📦 {prod} × {qty}\n"
            f"💰 {format_number(amt*qty)} تومان",
            reply_markup=InlineKeyboardMarkup(keyboard)
        )

    elif query.data.startswith("confirmdel_"):
        sale_id = int(query.data.split("_")[1])
        deleted = delete_sale(sale_id, user_id)
        if deleted:
            await query.edit_message_text("✅ فروش حذف شد.")
        else:
            await query.edit_message_text("❌ خطا در حذف فروش.")

    elif query.data == "canceldel":
        await query.edit_message_text("❌ حذف لغو شد.")

    elif query.data.startswith("edit_"):
        sale_id = int(query.data.split("_")[1])
        sale = get_sale_by_id(sale_id, user_id)
        if not sale:
            await query.edit_message_text("❌ این فروش پیدا نشد.")
            return
        context.user_data['edit_sale_id'] = sale_id
        keyboard = [[
            InlineKeyboardButton("📦 نام محصول", callback_data=f"editfield_product_{sale_id}"),
            InlineKeyboardButton("🔢 تعداد", callback_data=f"editfield_quantity_{sale_id}"),
        ], [
            InlineKeyboardButton("💵 قیمت واحد", callback_data=f"editfield_amount_{sale_id}"),
        ]]
        _, prod, amt, qty, date = sale
        await query.edit_message_text(
            f"✏️ کدوم بخش رو می‌خوای ویرایش کنی؟\n\n"
            f"📦 محصول: {prod}\n"
            f"🔢 تعداد: {qty}\n"
            f"💵 قیمت واحد: {format_number(amt)} تومان",
            reply_markup=InlineKeyboardMarkup(keyboard)
        )

    elif query.data.startswith("editfield_"):
        _, field, sale_id = query.data.split("_")
        context.user_data['edit_sale_id'] = int(sale_id)
        context.user_data['edit_field'] = field
        field_names = {"product": "نام محصول", "quantity": "تعداد", "amount": "قیمت واحد"}
        await query.edit_message_text(f"✏️ مقدار جدید برای «{field_names[field]}» رو بفرست:")
        context.user_data['awaiting_edit_value'] = True

async def process_edit_value(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    sale_id = context.user_data.get('edit_sale_id')
    field = context.user_data.get('edit_field')
    new_value_text = update.message.text.strip()

    if not sale_id or not field:
        await update.message.reply_text("❌ خطا، دوباره از منو امتحان کن.", reply_markup=main_keyboard())
        context.user_data.clear()
        return

    try:
        if field == "quantity":
            value = int(new_value_text)
            if value <= 0:
                raise ValueError
        elif field == "amount":
            value = float(new_value_text.replace(",", "").replace(" ", ""))
            if value < 0:
                raise ValueError
        else:  # product
            value = new_value_text
            if not value:
                raise ValueError

        updated = update_sale(sale_id, user_id, field, value)
        if updated:
            sale = get_sale_by_id(sale_id, user_id)
            _, prod, amt, qty, date = sale
            await update.message.reply_text(
                f"✅ ویرایش انجام شد!\n\n"
                f"📦 محصول: {prod}\n"
                f"🔢 تعداد: {qty}\n"
                f"💵 قیمت واحد: {format_number(amt)} تومان\n"
                f"💰 جمع: {format_number(amt*qty)} تومان",
                reply_markup=main_keyboard()
            )
        else:
            await update.message.reply_text("❌ خطا در ویرایش.", reply_markup=main_keyboard())

    except ValueError:
        await update.message.reply_text(
            "❌ مقدار نامعتبره. دوباره امتحان کن یا دوباره از منوی ویرایش شروع کن.",
            reply_markup=main_keyboard()
        )

    context.user_data.clear()

async def process_order_address(update: Update, context: ContextTypes.DEFAULT_TYPE):
    address_text = update.message.text.strip()
    order = context.user_data.get('last_order')

    if not order:
        await update.message.reply_text("❌ سفارشی پیدا نشد، دوباره از ثبت فروش شروع کن.", reply_markup=main_keyboard())
        context.user_data.clear()
        return

    # کل متن پیست‌شده عیناً به‌عنوان «مشخصات گیرنده» استفاده می‌شه، بدون پردازش یا تفکیک
    products_line = f"{order['product']}×{order['qty']}"
    context.user_data['pending_label_order'] = {
        'recipient_info': address_text,
        'products_line': products_line,
    }
    context.user_data.pop('awaiting_order_address', None)
    context.user_data.pop('last_order', None)

    keyboard = [[
        InlineKeyboardButton("📦 تیپاکس", callback_data="shipvia_tipax"),
        InlineKeyboardButton("📮 پست", callback_data="shipvia_post"),
        InlineKeyboardButton("🚚 چاپار", callback_data="shipvia_chapar"),
    ]]
    await update.message.reply_text(
        "🚚 روش ارسال چیه؟",
        reply_markup=InlineKeyboardMarkup(keyboard)
    )

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

    total = sum(s[2] * s[3] for s in sales)
    count = sum(s[3] for s in sales)

    text = f"☀️ گزارش امروز - {today}\n"
    text += "━━━━━━━━━━━━━━━━\n"

    for sid, p, a, q, d in sales:
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

    total = sum(s[2] * s[3] for s in sales)
    total_items = sum(s[3] for s in sales)

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

    total = sum(s[2] * s[3] for s in sales)
    total_items = sum(s[3] for s in sales)
    days_passed = (now - start).days + 1

    text = f"📈 گزارش ماهانه\n"
    text += f"📅 {start.strftime('%Y/%m')} | {days_passed} روز گذشته\n"
    text += "━━━━━━━━━━━━━━━━\n"
    text += f"🛒 کل سفارشات: {len(sales)} ({total_items} آیتم)\n"
    text += f"💰 درآمد کل: {format_number(total)} تومان\n"
    text += f"📌 میانگین روزانه: {format_number(total/days_passed)} تومان\n"
    text += f"📆 پیش‌بینی ماهانه: {format_number(total/days_passed*30)} تومان\n"
    estimated_profit = total * PROFIT_MARGIN
    text += f"💵 سود تقریبی (۳۰٪): {format_number(estimated_profit)} تومان\n"
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

    for sid, prod, amt, qty, date in sales:
        d = date.split(" ")[0]
        t = date.split(" ")[1][:5]
        text += f"📦 {prod} × {qty}\n"
        text += f"   💵 {format_number(amt*qty)} ت | 📅 {d} {t}\n\n"

    await update.message.reply_text(text)

# ===== خروجی اکسل =====
async def send_excel_export(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    now = datetime.now()
    start = now.replace(day=1, hour=0, minute=0, second=0)

    sales = get_sales_in_range(user_id,
        start.strftime("%Y-%m-%d 00:00:00"),
        now.strftime("%Y-%m-%d 23:59:59")
    )

    if not sales:
        await update.message.reply_text("📭 این ماه فروشی برای خروجی گرفتن وجود نداره.")
        return

    wb = Workbook()
    ws = wb.active
    ws.title = "فروش ماه"
    ws.sheet_view.rightToLeft = True

    headers = ["تاریخ", "نام محصول", "تعداد", "قیمت واحد", "مبلغ کل"]
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")

    total_all = 0
    for sale_id, prod, amt, qty, date in sorted(sales, key=lambda s: s[4]):
        d = date.split(" ")[0]
        row_total = amt * qty
        total_all += row_total
        ws.append([d, prod, qty, amt, row_total])

    ws.append(["", "", "", "جمع کل", total_all])
    last_row = ws.max_row
    ws[f"D{last_row}"].font = Font(bold=True)
    ws[f"E{last_row}"].font = Font(bold=True)

    for col in ["A", "B", "C", "D", "E"]:
        ws.column_dimensions[col].width = 18

    filename = f"/tmp/sales_{now.strftime('%Y_%m')}.xlsx"
    wb.save(filename)

    with open(filename, "rb") as f:
        await update.message.reply_document(
            document=f,
            filename=f"فروش_{now.strftime('%Y_%m')}.xlsx",
            caption=f"📥 خروجی اکسل فروش‌های {start.strftime('%Y/%m')} ✅"
        )


async def get_label_info(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text.strip()

    if context.user_data.get('standalone_label'):
        # مرحله ۱: مشخصات گیرنده دریافت شد، حالا محصولات رو بپرس
        context.user_data['standalone_recipient'] = text
        context.user_data.pop('standalone_label', None)
        context.user_data['standalone_awaiting_products'] = True
        await update.message.reply_text(
            "📦 حالا محصولات سفارش رو بفرست (مثال: تراول ماگ×۱, دستمال مرطوب×۲):"
        )
        return WAITING_LABEL_INFO

    if context.user_data.get('standalone_awaiting_products'):
        products_line = text
        recipient_info = context.user_data.pop('standalone_recipient', "")
        context.user_data.pop('standalone_awaiting_products', None)
        context.user_data['pending_label_order'] = {
            'recipient_info': recipient_info,
            'products_line': products_line,
        }
        keyboard = [[
            InlineKeyboardButton("📦 تیپاکس", callback_data="shipvia_tipax"),
            InlineKeyboardButton("📮 پست", callback_data="shipvia_post"),
            InlineKeyboardButton("🚚 چاپار", callback_data="shipvia_chapar"),
        ]]
        await update.message.reply_text(
            "🚚 روش ارسال چیه؟",
            reply_markup=InlineKeyboardMarkup(keyboard)
        )
        return ConversationHandler.END

    return ConversationHandler.END

def create_shipping_label(recipient_info, products_line, shipping_method=""):
    register_persian_font()
    use_font = FONT_NAME if FONT_REGISTERED else "Helvetica"

    filename = f"/tmp/label_{datetime.now().strftime('%Y%m%d%H%M%S')}.pdf"
    width, height = A5
    c = canvas.Canvas(filename, pagesize=A5)

    margin = 10 * mm
    y = height - margin

    def draw_rtl_line(text_str, font_size, y_pos):
        c.setFont(use_font, font_size)
        display_text = fa(text_str) if FONT_REGISTERED else text_str
        c.drawRightString(width - margin, y_pos, display_text)

    def wrap_text(text_str, max_chars):
        lines = []
        words = text_str.split(" ")
        current_line = ""
        for word in words:
            test_line = (current_line + " " + word).strip()
            if len(test_line) > max_chars:
                if current_line:
                    lines.append(current_line)
                current_line = word
            else:
                current_line = test_line
        if current_line:
            lines.append(current_line)
        return lines

    # کادر دور برگه
    c.setLineWidth(1.2)
    c.rect(4*mm, 4*mm, width - 8*mm, height - 8*mm)

    # لوگوی فروشگاه - گوشه سمت راست بالا
    logo_size = 16 * mm
    if os.path.exists(LOGO_PATH):
        try:
            logo = ImageReader(LOGO_PATH)
            c.drawImage(
                logo,
                width - margin - logo_size,
                height - margin - logo_size + 2*mm,
                width=logo_size,
                height=logo_size,
                preserveAspectRatio=True,
                mask='auto'
            )
        except Exception:
            pass

    # عنوان (پایین‌تر از لوگو، وسط‌چین)
    y -= (logo_size - 2*mm)
    c.setFont(use_font, 13)
    title_text = fa("فاکتور ارسال مرسوله") if FONT_REGISTERED else "فاکتور ارسال مرسوله"
    c.drawCentredString(width / 2, y, title_text)
    y -= 6*mm

    c.setLineWidth(0.5)
    c.line(margin, y, width - margin, y)
    y -= 6*mm

    # ===== بخش فرستنده (ثابت) =====
    draw_rtl_line("فرستنده:", 10.5, y)
    y -= 5.5*mm
    draw_rtl_line(f"{STORE_NAME} | {STORE_PHONE}", 9.5, y)
    y -= 5*mm
    for line in wrap_text(STORE_ADDRESS, 42):
        draw_rtl_line(line, 9, y)
        y -= 4.8*mm

    y -= 1.5*mm
    c.line(margin, y, width - margin, y)
    y -= 6*mm

    # ===== بخش گیرنده (مشتری) =====
    header = "گیرنده:"
    if shipping_method:
        header_full = f"گیرنده:        روش ارسال: {shipping_method}"
        draw_rtl_line(header_full, 10.5, y)
    else:
        draw_rtl_line(header, 10.5, y)
    y -= 6*mm

    for line in wrap_text(recipient_info, 40):
        draw_rtl_line(line, 11, y)
        y -= 6*mm

    y -= 2*mm
    c.line(margin, y, width - margin, y)
    y -= 6*mm

    # محصولات سفارش
    draw_rtl_line("اقلام سفارش:", 10.5, y)
    y -= 6*mm

    items = [p.strip() for p in products_line.split(",")]
    for item in items:
        draw_rtl_line(f"• {item}", 10, y)
        y -= 5.5*mm

    # پاورقی
    c.setFont("Helvetica", 7)
    c.drawCentredString(width / 2, 8*mm, datetime.now().strftime("%Y-%m-%d %H:%M"))

    c.save()
    return filename

# ===== اجرا =====
def main():
    init_db()
    register_persian_font()
    app = Application.builder().token(BOT_TOKEN).build()

    conv_handler = ConversationHandler(
        entry_points=[MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message)],
        states={
            WAITING_QUICK_SALE: [MessageHandler(filters.TEXT & ~filters.COMMAND, get_quick_sale)],
            WAITING_LABEL_INFO: [MessageHandler(filters.TEXT & ~filters.COMMAND, get_label_info)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
    )

    app.add_handler(CommandHandler("start", start))
    app.add_handler(conv_handler)
    app.add_handler(CallbackQueryHandler(handle_delete_callback))

    print("✅ ربات حسابداری شروع به کار کرد...")
    app.run_polling(allowed_updates=Update.ALL_TYPES)

if __name__ == "__main__":
    main()
