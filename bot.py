import asyncio
import logging
import os
from datetime import datetime
from io import BytesIO

from aiogram import Bot, Dispatcher, types, F
from aiogram.filters import Command
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.types import (
    InlineKeyboardMarkup,
    InlineKeyboardButton,
    BufferedInputFile
)

from docx import Document
from docx.shared import Pt


# ================== КОНФИГУРАЦИЯ ==================

TOKEN = os.getenv("BOT_TOKEN")  # ОБЯЗАТЕЛЬНО через env
if not TOKEN:
    raise RuntimeError("Не задан BOT_TOKEN в переменных окружения")

logging.basicConfig(level=logging.INFO)

bot = Bot(token=TOKEN)
dp = Dispatcher(storage=MemoryStorage())


# ================== FSM ==================

class ClaimStates(StatesGroup):
    choosing_marketplace = State()
    entering_reason = State()
    entering_full_name = State()
    entering_address = State()
    entering_order_num = State()
    entering_price = State()
    waiting_for_receipt = State()


# ================== ЮРИДИЧЕСКИЕ ДАННЫЕ ==================

LEGAL_BASE = {
    "WB": "ООО «Вайлдберриз», ИНН 7733545428, ОГРН 1067746062411. Адрес: 142181, МО, г. Подольск, д. Коледино, 6.",
    "OZON": "ООО «Интернет Решения», ИНН 7704217370, ОГРН 1027739244741. Адрес: 123112, г. Москва, Пресненская наб., 10.",
    "Yandex": "ООО «ЯНДЕКС», ИНН 7736207543, ОГРН 1027700229193. Адрес: 119021, г. Москва, ул. Льва Толстого, 16."
}


# ================== КЛАВИАТУРЫ ==================

def main_menu():
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="📝 Создать претензию", callback_data="create_claim")],
        [InlineKeyboardButton(text="📚 Правовой справочник", callback_data="legal_info")],
        [InlineKeyboardButton(text="📸 Распознать чек", callback_data="ocr_scan")]
    ])


def marketplace_kb():
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="Wildberries", callback_data="m_WB")],
        [InlineKeyboardButton(text="Ozon", callback_data="m_OZON")],
        [InlineKeyboardButton(text="Яндекс.Маркет", callback_data="m_Yandex")]
    ])


# ================== HANDLERS ==================

@dp.message(Command("start"))
async def start(message: types.Message):
    await message.answer(
        "👋 *Юрист-Бот по маркетплейсам*\n\n"
        "Помогаю составить юридически корректную досудебную претензию.",
        reply_markup=main_menu(),
        parse_mode="Markdown"
    )


@dp.callback_query(F.data == "create_claim")
async def start_claim(callback: types.CallbackQuery, state: FSMContext):
    await state.clear()
    await callback.message.answer(
        "Выберите маркетплейс:",
        reply_markup=marketplace_kb()
    )
    await state.set_state(ClaimStates.choosing_marketplace)
    await callback.answer()


@dp.callback_query(ClaimStates.choosing_marketplace, F.data.startswith("m_"))
async def choose_marketplace(callback: types.CallbackQuery, state: FSMContext):
    marketplace = callback.data.split("_")[1]
    await state.update_data(marketplace=marketplace)
    await callback.message.answer(
        "Кратко опишите проблему:"
    )
    await state.set_state(ClaimStates.entering_reason)
    await callback.answer()


@dp.message(ClaimStates.entering_reason)
async def enter_reason(message: types.Message, state: FSMContext):
    await state.update_data(reason=message.text)
    await message.answer("Введите ФИО полностью:")
    await state.set_state(ClaimStates.entering_full_name)


@dp.message(ClaimStates.entering_full_name)
async def enter_name(message: types.Message, state: FSMContext):
    await state.update_data(full_name=message.text)
    await message.answer("Введите почтовый адрес:")
    await state.set_state(ClaimStates.entering_address)


@dp.message(ClaimStates.entering_address)
async def enter_address(message: types.Message, state: FSMContext):
    await state.update_data(address=message.text)
    await message.answer("Введите номер заказа:")
    await state.set_state(ClaimStates.entering_order_num)


@dp.message(ClaimStates.entering_order_num)
async def enter_order(message: types.Message, state: FSMContext):
    await state.update_data(order_num=message.text)
    await message.answer("Введите сумму претензии (числом):")
    await state.set_state(ClaimStates.entering_price)


@dp.message(ClaimStates.entering_price)
async def enter_price(message: types.Message, state: FSMContext):
    try:
        price = float(message.text.replace(",", "."))
    except ValueError:
        await message.answer("❌ Введите сумму числом.")
        return

    await state.update_data(price=price)
    data = await state.get_data()

    await message.answer("⏳ Формирую документ...")

    doc_stream = create_docx(data)
    file = BufferedInputFile(
        doc_stream.getvalue(),
        filename=f"Pretenziya_{data['marketplace']}.docx"
    )

    await message.answer_document(
        file,
        caption="✅ Претензия готова. Распечатайте и отправьте заказным письмом."
    )
    await state.clear()


# ================== DOCX ==================

def create_docx(data: dict) -> BytesIO:
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Arial"
    style.font.size = Pt(12)

    doc.add_paragraph(f"Кому:\n{LEGAL_BASE[data['marketplace']]}\n").bold = True
    doc.add_paragraph(
        f"От:\n{data['full_name']}\n{data['address']}\n"
    )

    title = doc.add_paragraph("ДОСУДЕБНАЯ ПРЕТЕНЗИЯ")
    title.alignment = 1

    body = doc.add_paragraph()
    body.add_run(
        f"Я оформил заказ №{data['order_num']} на маркетплейсе "
        f"{data['marketplace']}. Возникла проблема: {data['reason']}.\n\n"
        f"Стоимость товара: {data['price']} руб.\n\n"
        "На основании ст. 18 и 22 Закона РФ «О защите прав потребителей» "
        "и ст. 309 ГК РФ\n\n"
    )

    body.add_run("ТРЕБУЮ:\n").bold = True
    body.add_run(
        f"Вернуть денежные средства в размере {data['price']} руб. "
        "в течение 10 календарных дней.\n\n"
    )

    body.add_run(
        f"Дата: {datetime.now().strftime('%d.%m.%Y')}   Подпись: ____________"
    )

    stream = BytesIO()
    doc.save(stream)
    stream.seek(0)
    return stream


# ================== СПРАВОЧНИК ==================

@dp.callback_query(F.data == "legal_info")
async def legal_info(callback: types.CallbackQuery):
    await callback.message.answer(
        "⚖️ *Правовая база*\n\n"
        "• ст. 18 ЗоЗПП\n"
        "• ст. 22 ЗоЗПП\n"
        "• ст. 309 ГК РФ\n\n"
        "Претензия обязательна перед судом.",
        parse_mode="Markdown"
    )
    await callback.answer()


# ================== OCR (ДЕМО) ==================

@dp.callback_query(F.data == "ocr_scan")
async def ocr_start(callback: types.CallbackQuery, state: FSMContext):
    await callback.message.answer(
        "📸 Отправьте фото чека (демо-режим OCR)."
    )
    await state.set_state(ClaimStates.waiting_for_receipt)
    await callback.answer()


@dp.message(ClaimStates.waiting_for_receipt, F.photo)
async def ocr_process(message: types.Message, state: FSMContext):
    await message.answer("🔍 Распознаю чек...")
    await asyncio.sleep(2)

    await message.answer(
        "✅ Чек получен (демо).\nНажмите «Создать претензию».",
        reply_markup=InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text="Создать претензию", callback_data="create_claim")]
        ])
    )
    await state.clear()

    # ================== ЗАПУСК ==================

async def main():
    print("🤖 Бот запущен и слушает Telegram")
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())



