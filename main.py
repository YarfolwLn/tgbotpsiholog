#Подключение библиотек
#pip install aiogram python-docx openpyxl

#Содержимое файла requirements.txt
#aiogram==3.13.0
#aiohttp==3.9.5
#openpyxl==3.1.5

#Установка зависимостей
#pip install -r requirements.txt

import asyncio
import logging
from aiogram import Bot, Dispatcher, types, F
from aiogram.filters import Command
from aiogram.types import ReplyKeyboardMarkup, KeyboardButton, ReplyKeyboardRemove
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.memory import MemoryStorage
import openpyxl
from openpyxl.styles import PatternFill
from openpyxl import load_workbook
from openpyxl.styles import NamedStyle
import os
from datetime import datetime, timedelta
from openpyxl.styles import Alignment

# Настройка логирования
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# Токен бота
BOT_TOKEN = os.getenv("TELEGRAM_TOKEN")

# Инициализация бота и диспетчера
bot = Bot(token=BOT_TOKEN)
storage = MemoryStorage()
dp = Dispatcher(storage=storage)

# Путь к файлу Excel
EXCEL_FILE = "appointments.xlsx"

# Создаем файл Excel если его нет
def init_excel_file():
    if not os.path.exists(EXCEL_FILE):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Записи"
        
        # Заголовки
        headers = ["Дата", "Время", "Имя пользователя", "Телеграм ID", "Телефон", "Ситуация", "Статус"]
        for col, header in enumerate(headers, 1):
            ws.cell(row=1, column=col, value=header)
        
        # Устанавливаем ширину колонок
        column_widths = [15, 10, 20, 15, 15, 30, 15]
        for col, width in enumerate(column_widths, 1):
            col_letter = openpyxl.utils.get_column_letter(col)
            ws.column_dimensions[col_letter].width = width
        
        # Создаем стили для ячеек
        date_style = NamedStyle(name="date_style", number_format='DD.MM.YYYY')
        time_style = NamedStyle(name="time_style", number_format='HH:MM')
        text_style = NamedStyle(name="text_style", number_format='@')  # Текстовый формат
        
        # Добавляем стили в книгу
        if 'date_style' not in wb.named_styles:
            wb.add_named_style(date_style)
        if 'time_style' not in wb.named_styles:
            wb.add_named_style(time_style)
        if 'text_style' not in wb.named_styles:
            wb.add_named_style(text_style)
        
        # Устанавливаем форматы для столбцов
        for row in range(2, 50):  # Устанавливаем для первых 50 строк
            ws.cell(row=row, column=1).style = 'date_style'    # Дата
            ws.cell(row=row, column=2).style = 'time_style'    # Время
            ws.cell(row=row, column=3).style = 'text_style'    # Имя
            ws.cell(row=row, column=4).style = 'text_style'    # Телеграм ID
            ws.cell(row=row, column=5).style = 'text_style'    # Телефон
            ws.cell(row=row, column=6).style = 'text_style'    # Ситуация
            ws.cell(row=row, column=7).style = 'text_style'    # Статус
        
        wb.save(EXCEL_FILE)
        logger.info("Excel файл создан с правильными форматами ячеек")

# Состояния FSM
class AppointmentState(StatesGroup):
    choosing_date = State()
    choosing_time = State()
    user_name = State()
    user_phone = State()
    user_situation = State()

# Класс для работы с Excel
class ExcelManager:
    def __init__(self, file_path):
        self.file_path = file_path
        self.red_fill = PatternFill(start_color="FFB6C1", end_color="FFB6C1", fill_type="solid")
    
    def book_appointment(self, date_str, time_str, username, user_id, phone, situation):
        """Просто записываем данные без проверки занятости"""
        try:
            wb = load_workbook(self.file_path)
            ws = wb.active
            
            # Находим первую свободную строку
            new_row = ws.max_row + 1
            
            # Записываем данные
            ws.cell(row=new_row, column=1, value=date_str)
            ws.cell(row=new_row, column=2, value=time_str)
            ws.cell(row=new_row, column=3, value=str(username))
            ws.cell(row=new_row, column=4, value=str(user_id))
            ws.cell(row=new_row, column=5, value=str(phone))
            ws.cell(row=new_row, column=6, value=str(situation))
            ws.cell(row=new_row, column=7, value="Ожидает подтверждения")  # Изменили статус
            
            # Красим строку для визуального выделения
            for col in range(1, 8):
                ws.cell(row=new_row, column=col).fill = self.red_fill
            
            wb.save(self.file_path)
            wb.close()
            return True
            
        except Exception as e:
            logger.error(f"Ошибка при записи в Excel: {e}")
            return False

    def get_user_appointments(self, user_id):
        """Получить записи пользователя"""
        try:
            wb = load_workbook(self.file_path)
            ws = wb.active
            
            user_appointments = []
            for row in range(2, ws.max_row + 1):
                user_id_cell = ws.cell(row=row, column=4)
                if user_id_cell.value == str(user_id):
                    date = ws.cell(row=row, column=1).value
                    time = ws.cell(row=row, column=2).value
                    situation = ws.cell(row=row, column=6).value
                    status = ws.cell(row=row, column=7).value
                    user_appointments.append({
                        'date': date,
                        'time': time,
                        'situation': situation,
                        'status': status
                    })
            
            wb.close()
            return user_appointments
            
        except Exception as e:
            logger.error(f"Ошибка при чтении записей пользователя: {e}")
            return []

# Инициализация менеджера Excel
excel_manager = ExcelManager(EXCEL_FILE)

# Клавиатуры
def get_main_keyboard():
    keyboard = ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="📅 Записаться на прием")],
            [KeyboardButton(text="📋 Мои записи")],
            [KeyboardButton(text="🆘 Помощь")]
        ],
        resize_keyboard=True,
        input_field_placeholder="Выберите действие..."
    )
    return keyboard

def get_exit_keyboard():
    """Клавиатура с кнопкой выхода"""
    return ReplyKeyboardMarkup(
        keyboard=[[KeyboardButton(text="🚪 Выход")]],
        resize_keyboard=True,
        one_time_keyboard=True
    )

# Функция для отправки уведомлений администратору
async def send_notification_to_admin(user_data, chosen_date_time, situation):
    """Отправка уведомления администратору о новой заявке"""
    try:
        admin_chat_id = os.getenv("ADMIN_ID")  # Замените на нужный ID администратора
        
        notification_text = (
            "🔔 **НОВАЯ ЗАЯВКА НА КОНСУЛЬТАЦИЮ**\n\n"
            f"📅 **Желаемая дата и время:** {chosen_date_time}\n"
            f"👤 **Имя клиента:** {user_data['user_name']}\n"
            f"📞 **Телефон:** {user_data['user_phone']}\n"
            f"🆔 **Telegram ID:** {user_data['user_id']}\n"
            f"📋 **Статус:** Ожидает подтверждения\n"
        )
        
        if situation:
            notification_text += f"📝 **Ситуация:** {situation}\n"
        
        notification_text += "\n⚠️ **Свяжитесь с клиентом для подтверждения записи**"
        
        await bot.send_message(chat_id=admin_chat_id, text=notification_text)
        logger.info(f"Уведомление отправлено администратору о новой заявке")
        
    except Exception as e:
        logger.error(f"Ошибка при отправке уведомления администратору: {e}")

# Обработчики команд
@dp.message(Command("start"))
async def cmd_start(message: types.Message, state: FSMContext):
    # Очищаем состояние, если пользователь был в процессе записи
    current_state = await state.get_state()
    if current_state:
        await state.clear()
        await message.answer(
            "❌ Процесс записи прерван. Возвращаемся в главное меню.",
            reply_markup=get_main_keyboard()
        )
    else:
        await message.answer(
            "👋 Добро пожаловать в бота по записи на приём!\n\n"
            "Я ваш виртуальный помощник. Я могу:\n"
            "• 📅 Записать вас на прием к психологу\n"
            "• 📋 Показать ваши активные записи\n\n"
            "Выберите действие из меню ниже:",
            reply_markup=get_main_keyboard()
        )

@dp.message(Command("help"))
async def cmd_help(message: types.Message, state: FSMContext):
    # Очищаем состояние, если пользователь был в процессе записи
    current_state = await state.get_state()
    if current_state:
        await state.clear()
        await message.answer(
            "❌ Процесс записи прерван. Возвращаемся в главное меню.",
            reply_markup=get_main_keyboard()
        )
    
    help_text = """
🆘 **Помощь по боту:**

📅 **Запись на прием:**
- Нажмите «📅 Записаться на прием»
- Введите желаемую дату и время
- Введите ваше имя, телефон и ситуацию
- Администратор свяжется с вами для подтверждения

📋 **Мои записи:**
- Просмотр ваших текущих заявок

🚪 **Управление:**
- «🚪 Выход» - прервать процесс записи
- «↩️ Назад» - вернуться в главное меню

Для начала работы нажмите /start
    """
    await message.answer(help_text)

@dp.message(F.text == "🆘 Помощь")
async def help_command(message: types.Message, state: FSMContext):
    await cmd_help(message, state)

@dp.message(F.text == "📅 Записаться на прием")
async def book_appointment(message: types.Message, state: FSMContext):
    await message.answer(
        "📅 Введите желаемую дату приема в формате ДД.ММ.ГГГГ (например, 25.12.2024):",
        reply_markup=get_exit_keyboard()
    )
    await state.set_state(AppointmentState.choosing_date)

# Обработка ввода даты
@dp.message(AppointmentState.choosing_date)
async def process_date_input(message: types.Message, state: FSMContext):
    if message.text == "🚪 Выход":
        await exit_process(message, state)
        return
        
    input_date = message.text.strip()
    await state.update_data(chosen_date=input_date)
    
    await message.answer(
        "⏰ Теперь введите желаемое время приема в формате ЧЧ:MM (например, 14:30):",
        reply_markup=get_exit_keyboard()
    )
    await state.set_state(AppointmentState.choosing_time)

# Обработка ввода времени
@dp.message(AppointmentState.choosing_time)
async def process_time_input(message: types.Message, state: FSMContext):
    if message.text == "🚪 Выход":
        await exit_process(message, state)
        return
        
    input_time = message.text.strip()
    await state.update_data(chosen_time=input_time)
    
    # Получаем выбранные дату и время
    user_data = await state.get_data()
    chosen_date = user_data['chosen_date']
    
    await message.answer(
        f"📅 Вы выбрали: {chosen_date} {input_time}\n\n"
        "Теперь введите ваше имя:",
        reply_markup=get_exit_keyboard()
    )
    await state.set_state(AppointmentState.user_name)

@dp.message(F.text == "📋 Мои записи")
async def my_appointments(message: types.Message):
    user_id = message.from_user.id
    try:
        appointments = excel_manager.get_user_appointments(user_id)
        
        if appointments:
            response = "📋 **Ваши заявки на консультацию:**\n\n"
            for i, appt in enumerate(appointments, 1):
                response += f"{i}. **Дата:** {appt['date']}\n"
                response += f"   **Время:** {appt['time']}\n"
                response += f"   **Статус:** {appt['status']}\n"
                if appt['situation']:
                    response += f"   **Ситуация:** {appt['situation']}\n"
                response += "\n"
            
            response += "📞 Администратор свяжется с вами для подтверждения записи."
            await message.answer(response)
        else:
            await message.answer(
                "📝 У вас пока нет активных заявок. "
                "Хотите оставить заявку на консультацию? Нажмите «📅 Записаться на прием»"
            )
            
    except Exception as e:
        logger.error(f"Ошибка при чтении записей: {e}")
        await message.answer("❌ Произошла ошибка при получении ваших записей. Попробуйте позже.")

@dp.message(F.text == "↩️ Назад")
async def back_to_main(message: types.Message, state: FSMContext):
    current_state = await state.get_state()
    if current_state:
        await state.clear()
    await message.answer("--Главное меню--\n"
"- «📅 Записаться на прием» - оставить заявку на консультацию\n"
"- «📋 Мои записи» - просмотр ваших заявок\n"
"Выберите действие из меню ниже:", reply_markup=get_main_keyboard())

@dp.message(F.text == "🚪 Выход")
async def exit_process(message: types.Message, state: FSMContext):
    current_state = await state.get_state()
    if current_state:
        await state.clear()
        await message.answer(
            "❌ Процесс записи прерван. Возвращаемся в главное меню.",
            reply_markup=get_main_keyboard()
        )
    else:
        await message.answer(
            "Вы уже в главном меню.",
            reply_markup=get_main_keyboard()
        )

# Обработка имени
@dp.message(AppointmentState.user_name)
async def process_name(message: types.Message, state: FSMContext):
    if message.text == "🚪 Выход":
        await exit_process(message, state)
        return
        
    if len(message.text.strip()) < 2:
        await message.answer("❌ Имя должно содержать хотя бы 2 символа. Пожалуйста, введите ваше имя:", reply_markup=get_exit_keyboard())
        return
        
    await state.update_data(user_name=message.text.strip())
    await message.answer("📞 Теперь введите ваш номер телефона (в любом формате):", reply_markup=get_exit_keyboard())
    await state.set_state(AppointmentState.user_phone)

# Обработка телефона
@dp.message(AppointmentState.user_phone)
async def process_phone(message: types.Message, state: FSMContext):
    if message.text == "🚪 Выход":
        await exit_process(message, state)
        return
        
    phone = message.text.strip()
    if len(phone) < 5:
        await message.answer("❌ Номер телефона слишком короткий. Пожалуйста, введите корректный номер:", reply_markup=get_exit_keyboard())
        return
        
    await state.update_data(user_phone=phone)
    await message.answer(
        "📝 Опишите кратко вашу ситуацию или проблему, с которой хотите обратиться "
        "(это поможет психологу лучше подготовиться к встрече):\n\n"
        "Если не хотите описывать, отправьте \"-\" или \"пропустить\"",
        reply_markup=get_exit_keyboard()
    )
    await state.set_state(AppointmentState.user_situation)

# Обработка ситуации и завершение записи
@dp.message(AppointmentState.user_situation)
async def process_situation(message: types.Message, state: FSMContext):
    if message.text == "🚪 Выход":
        await exit_process(message, state)
        return
        
    situation = message.text.strip()
    if situation.lower() in ["-", "пропустить", "нет", "не хочу"]:
        situation = ""
        
    user_data = await state.get_data()
    chosen_date = user_data['chosen_date']
    chosen_time = user_data['chosen_time']
    user_name = user_data['user_name']
    user_phone = user_data['user_phone']
    user_id = message.from_user.id
    
    # Добавляем user_id в user_data для уведомления
    user_data['user_id'] = user_id
    
    # Записываем в Excel
    success = excel_manager.book_appointment(chosen_date, chosen_time, user_name, user_id, user_phone, situation)
    
    if success:
        response = (
            f"✅ **Заявка успешно отправлена!**\n\n"
            f"📅 **Желаемая дата:** {chosen_date}\n"
            f"⏰ **Желаемое время:** {chosen_time}\n"
            f"👤 **Имя:** {user_name}\n"
            f"📞 **Телефон:** {user_phone}\n"
        )
        
        if situation:
            response += f"📝 **Ситуация:** {situation}\n\n"
        else:
            response += "\n"
            
        response += (
            "📞 Администратор свяжется с вами в ближайшее время для уточнения деталей и подтверждения записи.\n\n"
            "Спасибо за вашу заявку!"
        )
        
        await message.answer(response, reply_markup=get_main_keyboard())
        
        # Отправляем уведомление администратору
        await send_notification_to_admin(user_data, f"{chosen_date} {chosen_time}", situation)
        
    else:
        await message.answer(
            "❌ Произошла ошибка при отправке заявки. Пожалуйста, попробуйте позже.",
            reply_markup=get_main_keyboard()
        )
    
    await state.clear()

# Обработка обычных сообщений
@dp.message()
async def handle_other_messages(message: types.Message, state: FSMContext):
    if message.text.startswith('/'):
        return
        
    current_state = await state.get_state()
    if current_state:
        # Если пользователь в процессе записи, но ввел что-то не то
        await message.answer(
            "Пожалуйста, следуйте инструкциям процесса записи или нажмите «🚪 Выход» для отмены.",
            reply_markup=get_exit_keyboard()
        )
    else:
        await message.answer(
            "Я не совсем понимаю, что вы имеете в виду. "
            "Пожалуйста, используйте кнопки меню или команды для взаимодействия с ботом.",
            reply_markup=get_main_keyboard()
        )

# Основная функция
async def main():
    # Инициализируем Excel файл
    init_excel_file()
    
    logger.info("Бот запущен и готов к работе!")
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
