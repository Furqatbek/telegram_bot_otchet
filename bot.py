import os
import re
import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, Border, Side
from telegram import Update, ReplyKeyboardMarkup, ReplyKeyboardRemove
from telegram.ext import (
    Application,
    CommandHandler,
    ConversationHandler,
    MessageHandler,
    filters,
)

TYPE, AMOUNT, PAYMENT = range(3)

DATA_DIR = Path("data")
DATA_DIR.mkdir(exist_ok=True)

HEADERS = ["Kirim", "Chiqim", "Miqdor", "Naqd yoki Perechisleniya", "Sana"]

HEADER_FONT = Font(bold=True, size=12)
HEADER_ALIGNMENT = Alignment(horizontal="center")
THIN_BORDER = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin"),
)
COLUMN_WIDTHS = [15, 15, 18, 25, 20]


def get_user_file(user_id: int) -> Path:
    return DATA_DIR / f"{user_id}.xlsx"


def ensure_workbook(user_id: int) -> Path:
    path = get_user_file(user_id)
    if not path.exists():
        wb = Workbook()
        ws = wb.active
        ws.title = "Hisobot"
        for col_idx, header in enumerate(HEADERS, 1):
            cell = ws.cell(row=1, column=col_idx, value=header)
            cell.font = HEADER_FONT
            cell.alignment = HEADER_ALIGNMENT
            cell.border = THIN_BORDER
        for col_idx, width in enumerate(COLUMN_WIDTHS, 1):
            ws.column_dimensions[ws.cell(row=1, column=col_idx).column_letter].width = width
        wb.save(path)
    return path


def evaluate_math(expression: str) -> float | None:
    cleaned = expression.replace(" ", "").replace(",", ".")
    if not re.match(r"^[\d.+\-*/()]+$", cleaned):
        return None
    try:
        result = eval(cleaned, {"__builtins__": {}}, {})  # noqa: S307
        return float(result)
    except Exception:
        return None


def append_record(user_id: int, record_type: str, amount: float, payment: str) -> Path:
    path = ensure_workbook(user_id)
    wb = load_workbook(path)
    ws = wb.active

    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")

    kirim = amount if record_type == "Kirim" else ""
    chiqim = amount if record_type == "Chiqim" else ""

    row = [kirim, chiqim, amount, payment, now]
    next_row = ws.max_row + 1
    for col_idx, value in enumerate(row, 1):
        cell = ws.cell(row=next_row, column=col_idx, value=value)
        cell.border = THIN_BORDER
        if isinstance(value, float):
            cell.number_format = "#,##0.00"

    wb.save(path)
    return path


async def start(update: Update, context) -> int:
    keyboard = [["Kirim", "Chiqim"]]
    await update.message.reply_text(
        "Assalomu alaykum! Men sizning kirim-chiqim botingizman.\n\n"
        "Turini tanlang:",
        reply_markup=ReplyKeyboardMarkup(keyboard, one_time_keyboard=True, resize_keyboard=True),
    )
    return TYPE


async def choose_type(update: Update, context) -> int:
    text = update.message.text
    if text not in ("Kirim", "Chiqim"):
        keyboard = [["Kirim", "Chiqim"]]
        await update.message.reply_text(
            "Iltimos, 'Kirim' yoki 'Chiqim' ni tanlang:",
            reply_markup=ReplyKeyboardMarkup(keyboard, one_time_keyboard=True, resize_keyboard=True),
        )
        return TYPE

    context.user_data["type"] = text
    await update.message.reply_text(
        f"{text} tanlandi.\n\n"
        "Miqdorni kiriting (masalan: 50000 yoki 10000+25000+3000):",
        reply_markup=ReplyKeyboardRemove(),
    )
    return AMOUNT


async def enter_amount(update: Update, context) -> int:
    text = update.message.text.strip()
    result = evaluate_math(text)

    if result is None or result <= 0:
        await update.message.reply_text(
            "Noto'g'ri miqdor. Iltimos, raqam kiriting (masalan: 50000 yoki 10000+25000):"
        )
        return AMOUNT

    context.user_data["amount"] = result

    keyboard = [["Naqd", "Perechisleniya"]]
    display = f"{result:,.2f}".replace(",", " ")
    await update.message.reply_text(
        f"Miqdor: {display} so'm\n\nTo'lov turini tanlang:",
        reply_markup=ReplyKeyboardMarkup(keyboard, one_time_keyboard=True, resize_keyboard=True),
    )
    return PAYMENT


async def choose_payment(update: Update, context) -> int:
    text = update.message.text
    if text not in ("Naqd", "Perechisleniya"):
        keyboard = [["Naqd", "Perechisleniya"]]
        await update.message.reply_text(
            "Iltimos, 'Naqd' yoki 'Perechisleniya' ni tanlang:",
            reply_markup=ReplyKeyboardMarkup(keyboard, one_time_keyboard=True, resize_keyboard=True),
        )
        return PAYMENT

    record_type = context.user_data["type"]
    amount = context.user_data["amount"]
    user_id = update.effective_user.id

    path = append_record(user_id, record_type, amount, text)

    display = f"{amount:,.2f}".replace(",", " ")
    await update.message.reply_text(
        f"Saqlandi!\n"
        f"Turi: {record_type}\n"
        f"Miqdor: {display} so'm\n"
        f"To'lov: {text}\n\n"
        "Excel faylni yubormoqdaman...",
        reply_markup=ReplyKeyboardRemove(),
    )

    await update.message.reply_document(
        document=open(path, "rb"),
        filename=f"hisobot_{user_id}.xlsx",
        caption="Sizning hisobotingiz",
    )

    keyboard = [["Kirim", "Chiqim"]]
    await update.message.reply_text(
        "Yangi yozuv qo'shish uchun turini tanlang yoki /cancel bosing:",
        reply_markup=ReplyKeyboardMarkup(keyboard, one_time_keyboard=True, resize_keyboard=True),
    )
    return TYPE


async def cancel(update: Update, context) -> int:
    await update.message.reply_text(
        "Bekor qilindi. Qaytadan boshlash uchun /start bosing.",
        reply_markup=ReplyKeyboardRemove(),
    )
    return ConversationHandler.END


async def hisobot(update: Update, context) -> None:
    user_id = update.effective_user.id
    path = get_user_file(user_id)
    if not path.exists():
        await update.message.reply_text("Sizda hali yozuvlar yo'q. /start bosing.")
        return
    await update.message.reply_document(
        document=open(path, "rb"),
        filename=f"hisobot_{user_id}.xlsx",
        caption="Sizning hisobotingiz",
    )


def main():
    token = os.environ.get("BOT_TOKEN")
    if not token:
        raise SystemExit("BOT_TOKEN environment variable is not set")

    app = Application.builder().token(token).build()

    conv_handler = ConversationHandler(
        entry_points=[CommandHandler("start", start)],
        states={
            TYPE: [MessageHandler(filters.TEXT & ~filters.COMMAND, choose_type)],
            AMOUNT: [MessageHandler(filters.TEXT & ~filters.COMMAND, enter_amount)],
            PAYMENT: [MessageHandler(filters.TEXT & ~filters.COMMAND, choose_payment)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
    )

    app.add_handler(conv_handler)
    app.add_handler(CommandHandler("hisobot", hisobot))

    print("Bot ishga tushdi...")
    app.run_polling()


if __name__ == "__main__":
    main()
