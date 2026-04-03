import os
import tempfile
from pathlib import Path
from typing import Dict

from openpyxl import load_workbook
from PIL import Image, ImageDraw, ImageFont
from telegram import Update
from telegram.ext import Application, CommandHandler, ContextTypes, MessageHandler, filters

BOT_TOKEN = "8294093241:AAHKUedg20GgUOpMxAhKFUP3IocVhCKer0k"

TEMPLATE_PATH = Path("template.png")
FONT_PATH = Path("SBSansDisplay-Regular.ttf")

COLOR_DARK = (58, 68, 79, 255)
COLOR_POSITIVE = (124, 195, 42, 255)
COLOR_NEGATIVE = (79, 192, 242, 255)

FONT_DATE_SIZE = 32
FONT_VALUE_SIZE = 36
FONT_CHANGE_SIZE = 36

COORDS = {
    "date": (1115, 53),

    "imoex_value": (286, 239),
    "imoex_change": (286, 341),

    "rts_value": (518, 239),
    "rts_change": (518, 341),

    "brent_value": (739, 239),
    "brent_change": (739, 341),

    "gold_value": (966, 239),
    "gold_change": (966, 341),

    "cny_value": (286, 610),
    "cny_change": (286, 712),

    "usd_value": (518, 610),
    "usd_change": (518, 712),

    "eur_value": (739, 610),
    "eur_change": (739, 712),

    "aed_value": (966, 610),
    "aed_change": (966, 712),

    "bond2_value": (286, 984),
    "bond2_change": (286, 1088),

    "bond5_value": (518, 984),
    "bond5_change": (518, 1088),

    "bond10_value": (739, 984),
    "bond10_change": (739, 1088),

    "bond20_value": (966, 984),
    "bond20_change": (966, 1088),
}


def read_excel(xlsx_path: str) -> Dict[str, str]:
    wb = load_workbook(xlsx_path, data_only=True)
    ws = wb.active

    data = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0]:
            data[str(row[0])] = str(row[1]) if row[1] else ""
    return data


def get_color(text: str):
    if text.startswith("+"):
        return COLOR_POSITIVE
    if text.startswith("-"):
        return COLOR_NEGATIVE
    return COLOR_DARK


def render(template_path: Path, output_path: str, data: Dict[str, str]):
    image = Image.open(template_path).convert("RGBA")
    draw = ImageDraw.Draw(image)

    font_date = ImageFont.truetype(str(FONT_PATH), FONT_DATE_SIZE)
    font_value = ImageFont.truetype(str(FONT_PATH), FONT_VALUE_SIZE)
    font_change = ImageFont.truetype(str(FONT_PATH), FONT_CHANGE_SIZE)

    draw.text(COORDS["date"], data.get("date", ""), font=font_date, fill=COLOR_DARK)

    for key in COORDS:
        if "value" in key:
            draw.text(COORDS[key], data.get(key, ""), font=font_value, fill=COLOR_DARK)

        if "change" in key:
            val = data.get(key, "")
            draw.text(COORDS[key], val, font=font_change, fill=get_color(val))

    image.save(output_path)


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Отправь Excel файл .xlsx")


async def handle(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document

    if not doc.file_name.endswith(".xlsx"):
        await update.message.reply_text("Нужен .xlsx файл")
        return

    with tempfile.TemporaryDirectory() as tmp:
        file_path = os.path.join(tmp, doc.file_name)
        result_path = os.path.join(tmp, "result.png")

        file = await context.bot.get_file(doc.file_id)
        await file.download_to_drive(file_path)

        data = read_excel(file_path)
        render(TEMPLATE_PATH, result_path, data)

        with open(result_path, "rb") as img:
            await update.message.reply_photo(img)


def main():
    if not TEMPLATE_PATH.exists():
        raise Exception("Нет template.png")

    if not FONT_PATH.exists():
        raise Exception(f"Нет шрифта {FONT_PATH}")

    app = Application.builder().token(BOT_TOKEN).build()

    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.Document.ALL, handle))

    print("Бот запущен 🚀")
    app.run_polling()


if __name__ == "__main__":
    main()
