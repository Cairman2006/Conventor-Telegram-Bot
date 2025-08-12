# Conventor-Telegram-Bot
import asyncio
import os
import uuid
import platform
import subprocess
import pypandoc
from aiogram import Bot, Dispatcher, types, F
from aiogram.types import Message, CallbackQuery, FSInputFile
from aiogram.filters import CommandStart
from aiogram.utils.keyboard import InlineKeyboardBuilder

from docx2pdf import convert  # Только Windows/macOS
from pdf2docx import Converter  # Для PDF → Word

TOKEN = "TOKEN file"  # ← убран пробел

bot = Bot(token=TOKEN)
dp = Dispatcher()

# Состояния пользователя
user_state = {}

# Клавиатура выбора
def get_conversion_keyboard():
    kb = InlineKeyboardBuilder()
    kb.button(text="📄 Word → PDF", callback_data="word_to_pdf")
    kb.button(text="📑 PDF → Word", callback_data="pdf_to_word")
    kb.button(text="📝 Word → TXT", callback_data="word_to_txt")
    kb.button(text="🌐 Word → HTML", callback_data="word_to_html")
    kb.button(text="🧾 Word → ODT", callback_data="word_to_odt")
    kb.adjust(1)
    return kb.as_markup()

@dp.message(CommandStart())
async def start_handler(message: Message):
    await message.answer(
        "Привет! Что ты хочешь сделать?",
        reply_markup=get_conversion_keyboard()
    )

@dp.callback_query()
async def process_callback(callback: CallbackQuery):
    user_id = callback.from_user.id
    data = callback.data
    user_state[user_id] = data

    action_text = {
        "word_to_pdf": "Пришли .docx файл — я сделаю PDF.",
        "pdf_to_word": "Пришли .pdf файл — я сделаю Word.",
        "word_to_txt": "Пришли .docx файл — я сделаю TXT.",
        "word_to_html": "Пришли .docx файл — я сделаю HTML.",
        "word_to_odt": "Пришли .docx файл — я сделаю ODT."
    }

    await callback.message.answer(action_text.get(data, "Формат не поддерживается."))
    await callback.answer()

@dp.message(F.document)
async def handle_document(message: Message):
    user_id = message.from_user.id
    state = user_state.get(user_id)

    if not state:
        await message.answer("Сначала выбери действие через /start.")
        return

    file = message.document
    file_name = file.file_name
    uid = str(uuid.uuid4())
    input_path = f"{uid}_{file_name}"
    output_path = ""

    await bot.download(file, destination=input_path)

    try:
        if state == "word_to_pdf":
            if not file_name.endswith(".docx"):
                await message.answer("Это не .docx файл.")
                return
            output_path = f"{uid}.pdf"
            if platform.system() in ["Windows", "Darwin"]:
                convert(input_path, output_path)
            else:
                subprocess.run([
                    "libreoffice", "--headless", "--convert-to", "pdf", input_path, "--outdir", "."
                ], check=True)

        elif state == "pdf_to_word":
            if not file_name.endswith(".pdf"):
                await message.answer("Это не .pdf файл.")
                return
            output_path = f"{uid}.docx"
            cv = Converter(input_path)
            cv.convert(output_path, start=0, end=None)
            cv.close()

        elif state.startswith("word_to_"):
            if not file_name.endswith(".docx"):
                await message.answer("Это не .docx файл.")
                return

            format_map = {
                "word_to_txt": "plain",
                "word_to_html": "html",
                "word_to_odt": "odt"
            }

            to_format = format_map.get(state)
            ext = "txt" if to_format == "plain" else to_format
            output_path = f"{uid}_converted.{ext}"

            pypandoc.convert_file(
                input_path,
                to_format,
                outputfile=output_path,
                extra_args=["--standalone"]
            )

        else:
            await message.answer("Неизвестный режим.")
            return

        await message.answer_document(FSInputFile(output_path))

    except Exception as e:
        await message.answer(f"❌ Ошибка при обработке файла:\n`{e}`", parse_mode="Markdown")
    finally:
        for f in [input_path, output_path]:
            if f and os.path.exists(f):
                os.remove(f)

async def main():
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
