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
        
        # Добавляем примеры дат и времени
        sample_appointments = [
            ("15.12.2024", "10:00"),
            ("15.12.2024", "14:00"), 
            ("16.12.2024", "11:00"),
            ("16.12.2024", "15:00"),
            ("17.12.2024", "10:00"),
            ("17.12.2024", "16:00"),
            ("18.12.2024", "12:00"),
            ("18.12.2024", "17:00")
        ]
        
        # Подсказка для корректного заполнения таблицы
        ws.merge_cells('L1:P3')
        ws['L1'] = "___Правило,  чтобы все корректно работало___\n Все данные в таблице указываются в следующем формате: '15.12.2024"
        ws['L1'].alignment = Alignment(wrap_text=True, vertical='top', horizontal='left')
        

        for i, (date, time) in enumerate(sample_appointments, start=2):
            ws.cell(row=i, column=1, value=date)
            ws.cell(row=i, column=2, value=time)
            ws.cell(row=i, column=7, value="Свободно")
        
        wb.save(EXCEL_FILE)
        logger.info("Excel файл создан с правильными форматами ячеек")

# Состояния FSM
class AppointmentState(StatesGroup):
    choosing_date = State()
    user_name = State()
    user_phone = State()
    user_situation = State()

# Класс для работы с Excel
class ExcelManager:
    def __init__(self, file_path):
        self.file_path = file_path
        self.red_fill = PatternFill(start_color="FFB6C1", end_color="FFB6C1", fill_type="solid")
    
    def get_available_dates(self):
        """Получить список свободных дат и времени в формате 'Дата Время'"""
        try:
            wb = load_workbook(self.file_path)
            ws = wb.active
            dates = []
            
            for row in range(2, ws.max_row + 1):
                status_cell = ws.cell(row=row, column=7)
                if status_cell.value == "Свободно":
                    date_cell = ws.cell(row=row, column=1)
                    time_cell = ws.cell(row=row, column=2)
                    if date_cell.value and time_cell.value:
                        # Получаем значения как строки
                        date_str = str(date_cell.value)
                        time_str = str(time_cell.value)
                        dates.append(f"{date_str} {time_str}")
            
            wb.close()
            return dates
        except Exception as e:
            logger.error(f"Ошибка при чтении Excel: {e}")   
            return []
    
    def get_available_dates_with_times(self):
        """Получить словарь с датами и доступным временем"""
        try:
            wb = load_workbook(self.file_path)
            ws = wb.active
            dates_dict = {}
            
            for row in range(2, ws.max_row + 1):
                status_cell = ws.cell(row=row, column=7)
                date_cell = ws.cell(row=row, column=1)
                time_cell = ws.cell(row=row, column=2)
                
                if (status_cell.value == "Свободно" and 
                    date_cell.value and time_cell.value):
                    date_str = str(date_cell.value)
                    time_str = str(time_cell.value)
                    
                    if date_str not in dates_dict:
                        dates_dict[date_str] = []
                    dates_dict[date_str].append(time_str)
            
            wb.close()
            return dates_dict
        except Exception as e:
            logger.error(f"Ошибка при чтении Excel: {e}")   
            return {}
    
    def book_appointment(self, date_time_str, username, user_id, phone, situation):
        """Забронировать время"""
        try:
            # Разделяем строку на дату и время
            date_str, time_str = date_time_str.split(" ", 1)
            
            wb = load_workbook(self.file_path)
            ws = wb.active
            success = False
            
            for row in range(2, ws.max_row + 1):
                date_cell = ws.cell(row=row, column=1)
                time_cell = ws.cell(row=row, column=2)
                status_cell = ws.cell(row=row, column=7)
                
                if (date_cell.value and time_cell.value and
                    str(date_cell.value) == date_str and 
                    str(time_cell.value) == time_str and 
                    status_cell.value == "Свободно"):
                    
                    ws.cell(row=row, column=3, value=str(username))
                    ws.cell(row=row, column=4, value=str(user_id))
                    ws.cell(row=row, column=5, value=str(phone))
                    ws.cell(row=row, column=6, value=str(situation))
                    ws.cell(row=row, column=7, value="Забронировано")
                    
                    # Красим всю строку в красный
                    for col in range(1, 8):
                        ws.cell(row=row, column=col).fill = self.red_fill
                    
                    success = True
                    break
            
            if success:
                wb.save(self.file_path)
            
            wb.close()
            return success
            
        except Exception as e:
            logger.error(f"Ошибка при записи в Excel: {e}")
            return False

    def add_new_slots(self, date, times):
        """Добавить новые временные слоты для даты"""
        try:
            wb = load_workbook(self.file_path)
            ws = wb.active
            
            # Находим последнюю строку
            last_row = ws.max_row + 1
            
            for time_slot in times:
                # Проверяем, существует ли уже такая запись
                exists = False
                for row in range(2, last_row):
                    date_cell = ws.cell(row=row, column=1)
                    time_cell = ws.cell(row=row, column=2)
                    if (str(date_cell.value) == date and 
                        str(time_cell.value) == time_slot):
                        exists = True
                        break
                
                if not exists:
                    ws.cell(row=last_row, column=1, value=date)
                    ws.cell(row=last_row, column=2, value=time_slot)
                    ws.cell(row=last_row, column=7, value="Свободно")
                    last_row += 1
            
            wb.save(self.file_path)
            wb.close()
            return True
            
        except Exception as e:
            logger.error(f"Ошибка при добавлении слотов: {e}")
            return False

# Инициализация менеджера Excel
excel_manager = ExcelManager(EXCEL_FILE)

# Клавиатуры
def get_main_keyboard():
    keyboard = ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="📅 Записаться на прием")],
            [KeyboardButton(text="📋 Мои записи"), KeyboardButton(text="📅 Свободные даты")],
            [KeyboardButton(text="🆘 Помощь")]
        ],
        resize_keyboard=True,
        input_field_placeholder="Выберите действие..."
    )
    return keyboard

def get_dates_keyboard():
    """Клавиатура с доступными датами и временем"""
    dates = excel_manager.get_available_dates()
    keyboard = []
    
    # Группируем даты по 2 в строке
    for i in range(0, len(dates), 2):
        row = [KeyboardButton(text=dates[i])]
        if i + 1 < len(dates):
            row.append(KeyboardButton(text=dates[i + 1]))
        keyboard.append(row)
    
    keyboard.append([KeyboardButton(text="↩️ Назад")])
    return ReplyKeyboardMarkup(keyboard=keyboard, resize_keyboard=True)

def get_exit_keyboard():
    """Клавиатура с кнопкой выхода"""
    return ReplyKeyboardMarkup(
        keyboard=[[KeyboardButton(text="🚪 Выход")]],
        resize_keyboard=True,
        one_time_keyboard=True
    )

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
            "• 📋 Показать ваши активные записи\n"
            "• 📅 Показать свободные даты для записи\n\n"
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
- Выберите удобную дату и время из списка
- Введите ваше имя, ситуацию и телефон
- В любой момент можно нажать «🚪 Выход» для отмены записи

ℹ️ **Информация:**
- «📅 Свободные даты» - доступное время для записи

📋 **Управление:**
- «📋 Мои записи» - просмотр ваших записей
- «↩️ Назад» - вернуться в главное меню
- «🚪 Выход» - прервать процесс записи

Для начала работы нажмите /start
    """
    await message.answer(help_text)

@dp.message(F.text == "🆘 Помощь")
async def help_command(message: types.Message, state: FSMContext):
    await cmd_help(message, state)

@dp.message(F.text == "📅 Записаться на прием")
async def book_appointment(message: types.Message, state: FSMContext):
    dates = excel_manager.get_available_dates()
    if not dates:
        await message.answer(
            "😔 К сожалению, все время занято. "
            "Новые даты появляются регулярно - проверяйте позже или напишите нам для уточнения свободных окон."
        )
        return
    
    await message.answer(
        "📅 Выберите удобную дату и время из доступных:\n\n" +
        "\n".join([f"• {date}" for date in dates]),
        reply_markup=get_dates_keyboard()
    )
    await state.set_state(AppointmentState.choosing_date)

@dp.message(F.text == "📅 Свободные даты")
async def show_available_dates(message: types.Message):
    dates_dict = excel_manager.get_available_dates_with_times()
    if dates_dict:
        response = "📅 **Свободные даты и время для записи:**\n\n"
        for date, times in dates_dict.items():
            response += f"**{date}:**\n"
            response += "\n".join([f"• {time}" for time in sorted(times)]) + "\n\n"
        response += "Для записи нажмите «📅 Записаться на прием»"
        await message.answer(response)
    else:
        await message.answer(
            "⏳ На данный момент все время занято. "
            "Новые даты появятся в ближайшее время - проверяйте регулярно!"
        )

@dp.message(F.text == "📋 Мои записи")
async def my_appointments(message: types.Message):
    user_id = message.from_user.id
    try:
        wb = load_workbook(EXCEL_FILE)
        ws = wb.active
        
        user_appointments = []
        for row in range(2, ws.max_row + 1):
            user_id_cell = ws.cell(row=row, column=4)
            if user_id_cell.value == str(user_id):
                date = ws.cell(row=row, column=1).value
                time = ws.cell(row=row, column=2).value
                situation = ws.cell(row=row, column=6).value
                status = ws.cell(row=row, column=7).value
                situation_text = f"\n   📝 Ситуация: {situation}" if situation else ""
                user_appointments.append(f"✅ {date} {time} - {status}{situation_text}")
        
        wb.close()
        
        if user_appointments:
            await message.answer("📋 **Ваши записи:**\n\n" + "\n".join(user_appointments))
        else:
            await message.answer(
                "📝 У вас пока нет активных записей. "
                "Хотите записаться на прием? Нажмите «📅 Записаться на прием»"
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
"- «📅 Записаться на прием» - запись на приём\n"
"- «📅 Свободные даты» - доступное время для записи\n"
"- «📋 Мои записи» - просмотр ваших записей\n"
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

# Обработка выбора даты
@dp.message(AppointmentState.choosing_date)
async def process_date_choice(message: types.Message, state: FSMContext):
    if message.text == "↩️ Назад":
        await back_to_main(message, state)
        return
        
    if message.text == "🚪 Выход":
        await exit_process(message, state)
        return
        
    chosen_date_time = message.text
    dates = excel_manager.get_available_dates()
    
    if chosen_date_time not in dates:
        await message.answer("❌ Пожалуйста, выберите дату и время из предложенных вариантов кнопками ниже.")
        return
    
    await state.update_data(chosen_date=chosen_date_time)
    await message.answer(
        f"📅 Вы выбрали: {chosen_date_time}\n\n"
        "Теперь введите ваше имя:",
        reply_markup=get_exit_keyboard()
    )
    await state.set_state(AppointmentState.user_name)

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

# Обработка телефона и завершение записи
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

# Добавьте эту функцию для отправки уведомлений
async def send_notification_to_admin(user_data, chosen_date_time, situation):
    """Отправка уведомления администратору о новой записи"""
    try:
        admin_chat_id = os.getenv("ADMIN_ID")  # Замените на нужный ID или username
        
        notification_text = (
            "🔔 **НОВАЯ ЗАПИСЬ НА КОНСУЛЬТАЦИЮ**\n\n"
            f"📅 **Дата и время:** {chosen_date_time}\n"
            f"👤 **Имя клиента:** {user_data['user_name']}\n"
            f"📞 **Телефон:** {user_data['user_phone']}\n"
            f"🆔 **Telegram ID:** {user_data['user_id']}\n"
        )
        
        if situation:
            notification_text += f"📝 **Ситуация:** {situation}\n"
        
        await bot.send_message(chat_id=admin_chat_id, text=notification_text)
        logger.info(f"Уведомление отправлено администратору о записи {chosen_date_time}")
        
    except Exception as e:
        logger.error(f"Ошибка при отправке уведомления администратору: {e}")

# Модифицируйте функцию process_situation
@dp.message(AppointmentState.user_situation)
async def process_situation(message: types.Message, state: FSMContext):
    if message.text == "🚪 Выход":
        await exit_process(message, state)
        return
        
    situation = message.text.strip()
    if situation.lower() in ["-", "пропустить", "нет", "не хочу"]:
        situation = ""
        
    user_data = await state.get_data()
    chosen_date_time = user_data['chosen_date']
    user_name = user_data['user_name']
    user_phone = user_data['user_phone']
    user_id = message.from_user.id
    
    # Добавляем user_id в user_data для уведомления
    user_data['user_id'] = user_id
    
    # Записываем в Excel
    success = excel_manager.book_appointment(chosen_date_time, user_name, user_id, user_phone, situation)
    
    if success:
        response = (
            f"🎉 **Запись успешно оформлена!**\n\n"
            f"📅 **Дата и время:** {chosen_date_time}\n"
            f"👤 **Имя:** {user_name}\n"
            f"📞 **Телефон:** {user_phone}\n"
        )
        
        if situation:
            response += f"📝 **Ситуация:** {situation}\n\n"
        else:
            response += "\n"
            
        response += (
            f"Мы ждем вас на консультации! За день до приема напомним о встрече.\n\n"
            f"Если у вас возникли вопросы - напишите нам."
        )
        
        await message.answer(response, reply_markup=get_main_keyboard())
        
        # Отправляем уведомление администратору
        await send_notification_to_admin(user_data, chosen_date_time, situation)
        
    else:
        await message.answer(
            "❌ К сожалению, это время уже занято. Пожалуйста, выберите другое время из списка.",
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



