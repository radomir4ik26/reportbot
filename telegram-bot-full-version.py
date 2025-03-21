import logging 
import asyncio
from aiogram import Bot, Dispatcher, types, F
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.types import ReplyKeyboardMarkup, KeyboardButton, ReplyKeyboardRemove
from docx import Document
import os

# Налаштування бота
API_TOKEN = "8035646713:AAGaYfc6NcmAHR0iseSNu7Vcs2N6tOodlXI"
bot = Bot(token=API_TOKEN)
storage = MemoryStorage()

# Визначення шляхів до файлів
template_path = "report_template.docx"
names_path = "names.txt"
reasons_path = "reasons.txt"

# Список доступних звань
RANKS = ["солдат", "старший солдат", "молодший сержант", "сержант"]

# Функція для завантаження списку курсантів
def load_names():
    try:
        with open(names_path, 'r', encoding='utf-8') as file:
            return [line.strip() for line in file if line.strip()]
    except Exception as e:
        logging.error(f"Помилка при завантаженні списку курсантів: {e}")
        return []

# Функція для завантаження причин рапортів
def load_reasons():
    try:
        with open(reasons_path, 'r', encoding='utf-8') as file:
            return [line.strip() for line in file if line.strip()]
    except Exception as e:
        logging.error(f"Помилка при завантаженні причин рапортів: {e}")
        return []

# Визначення стандартного командувача
DEFAULT_COMMANDER = "Командиру навчальної групи №305"

# Визначення станів для FSM
class ReportStates(StatesGroup):
    waiting_for_name = State()
    waiting_for_rank = State()
    waiting_for_time_date = State()
    waiting_for_reason = State()
    waiting_for_address = State()
    waiting_for_report_date = State()

# Ініціалізація Dispatcher
dp = Dispatcher(storage=storage)

# Налаштування логування
logging.basicConfig(level=logging.INFO)

# Обробник команди /start
@dp.message(F.text == "/start")
async def start_handler(message: types.Message):
    markup = ReplyKeyboardMarkup(
        keyboard=[[KeyboardButton(text="Створити рапорт")]],
        resize_keyboard=True
    )
    await message.answer("Вітаю! Я допоможу вам створити рапорт. Натисніть кнопку, щоб почати.", reply_markup=markup)

# Обробник кнопки "Створити рапорт"
@dp.message(F.text == "Створити рапорт")
async def create_report(message: types.Message, state: FSMContext):
    names = load_names()
    
    if not names:
        await message.answer("Помилка: Не вдалося завантажити список курсантів. Спробуйте пізніше.")
        return
    
    # Створюємо клавіатуру з іменами курсантів
    keyboard = []
    for name in names:
        keyboard.append([KeyboardButton(text=name)])
    
    markup = ReplyKeyboardMarkup(
        keyboard=keyboard,
        resize_keyboard=True
    )
    
    await message.answer("Оберіть своє ім'я зі списку:", reply_markup=markup)
    await state.set_state(ReportStates.waiting_for_name)

# Обробник вибору імені курсанта
@dp.message(ReportStates.waiting_for_name)
async def select_name(message: types.Message, state: FSMContext):
    name = message.text
    names = load_names()
    
    if name not in names and names:
        await message.answer("Будь ласка, оберіть ім'я зі списку.")
        return
    
    await state.update_data(name=name)
    await state.update_data(commander=DEFAULT_COMMANDER)  # Використовуємо стандартного командувача
    
    # Створюємо клавіатуру для вибору звання
    keyboard = []
    for rank in RANKS:
        keyboard.append([KeyboardButton(text=rank)])
    
    markup = ReplyKeyboardMarkup(
        keyboard=keyboard,
        resize_keyboard=True
    )
    
    await message.answer("Оберіть ваше звання:", reply_markup=markup)
    await state.set_state(ReportStates.waiting_for_rank)

# Обробник введення звання
@dp.message(ReportStates.waiting_for_rank)
async def handle_rank(message: types.Message, state: FSMContext):
    rank = message.text
    
    # Перевіряємо, чи є звання в списку
    if rank not in RANKS:
        await message.answer("Будь ласка, оберіть звання з запропонованого списку.")
        return
    
    await state.update_data(rank=rank)
    
    await message.answer("Введіть час і дату рапорту (наприклад, 08:00 21.12.2024):")
    await state.set_state(ReportStates.waiting_for_time_date)

# Обробник введення часу і дати
@dp.message(ReportStates.waiting_for_time_date)
async def handle_time_date(message: types.Message, state: FSMContext):
    time_date = message.text
    await state.update_data(time_date=time_date)
    
    reasons = load_reasons()
    
    if not reasons:
        # Якщо не вдалося завантажити причини, просимо користувача ввести причину вручну
        await message.answer("Введіть причину рапорту:")
    else:
        # Створюємо клавіатуру з причинами
        keyboard = []
        for reason in reasons:
            keyboard.append([KeyboardButton(text=reason)])
        
        markup = ReplyKeyboardMarkup(
            keyboard=keyboard,
            resize_keyboard=True
        )
        
        await message.answer("Оберіть причину рапорту:", reply_markup=markup)
    
    await state.set_state(ReportStates.waiting_for_reason)

# Обробник вибору причини
@dp.message(ReportStates.waiting_for_reason)
async def handle_reason(message: types.Message, state: FSMContext):
    reason = message.text
    await state.update_data(reason=reason)
    
    await message.answer("Введіть адресу звільнення:", reply_markup=ReplyKeyboardRemove())
    await state.set_state(ReportStates.waiting_for_address)

# Обробник введення адреси
@dp.message(ReportStates.waiting_for_address)
async def handle_address(message: types.Message, state: FSMContext):
    address = message.text
    await state.update_data(address=address)
    
    await message.answer("Введіть дату підпису рапорту:")
    await state.set_state(ReportStates.waiting_for_report_date)

# Обробник введення дати підпису та створення рапорту
@dp.message(ReportStates.waiting_for_report_date)
async def handle_report_date(message: types.Message, state: FSMContext):
    report_date = message.text
    await state.update_data(report_date=report_date)
    
    # Отримуємо всі дані зі стану
    user_data = await state.get_data()
    
    await message.answer("Створення рапорту...")
    
    try:
        # Генерація рапорту
        generated_file = generate_report(
            template_path, 
            user_data['commander'],
            user_data['name'],
            user_data['rank'], 
            user_data['time_date'], 
            user_data['reason'], 
            user_data['address'], 
            user_data['report_date']
        )
        
        # Надсилаємо згенерований файл користувачу
        await bot.send_document(message.chat.id, types.FSInputFile(generated_file))
        
        # Видаляємо файл після надсилання
        os.remove(generated_file)
        
        # Очищаємо стан
        await state.clear()
        
        # Показуємо кнопку для створення нового рапорту
        markup = ReplyKeyboardMarkup(
            keyboard=[[KeyboardButton(text="Створити рапорт")]],
            resize_keyboard=True
        )
        await message.answer("Рапорт створено і відправлено! Бажаєте створити ще один?", reply_markup=markup)
    
    except Exception as e:
        logging.error(f"Помилка при створенні рапорту: {e}")
        await message.answer(f"Виникла помилка при створенні рапорту: {e}")
        await state.clear()

def generate_report(template_path, commander, name, rank, time_date, reason, address, report_date):
    try:
        doc = Document(template_path)
        
        # Розбираємо компоненти імені
        name_parts = name.split()
        if len(name_parts) >= 3:  # Формат: Прізвище Ім'я По-батькові
            last_name = name_parts[0]  # Прізвище в давальному відмінку
            first_name = name_parts[1]  # Ім'я
            middle_name = name_parts[2]  # По-батькові
            
            # Перетворюємо прізвище з давального відмінка (кому?) на родовий (кого?)
            # Типові закінчення -у/-ю в давальному -> -а/-я в родовому
            if last_name.endswith("у"):
                last_name_genitive = last_name[:-1] + "а"
            elif last_name.endswith("ю"):
                last_name_genitive = last_name[:-1] + "я"
            else:
                # Для жіночих прізвищ та інших випадків
                last_name_genitive = last_name
            
            # Змінений формат: Ім'я ПРІЗВИЩЕ (наприклад, Радомир ВІНІЧЕНКО)
            formal_name = f"{first_name} {last_name.upper()}"
            
            # Формат для згадування: ПРІЗВИЩА І.П. (у родовому відмінку)
            name_with_uppercase_lastname = f"{last_name_genitive.upper()} {first_name[0].upper()}.{middle_name[0].upper()}."
            
            # Повне ім'я з формальним відмінком
            full_name = f"{rank} {name}"
            
        elif len(name_parts) == 2:  # Формат: Прізвище Ім'я
            last_name = name_parts[0]  # Прізвище
            first_name = name_parts[1]  # Ім'я
            
            # Перетворюємо прізвище з давального відмінка на родовий
            if last_name.endswith("у"):
                last_name_genitive = last_name[:-1] + "а"
            elif last_name.endswith("ю"):
                last_name_genitive = last_name[:-1] + "я"
            else:
                # Для жіночих прізвищ та інших випадків
                last_name_genitive = last_name
            
            # Змінений формат: Ім'я ПРІЗВИЩЕ (наприклад, Радомир ВІНІЧЕНКО)
            formal_name = f"{first_name} {last_name.upper()}"
            
            # Формат для згадування: ПРІЗВИЩА І. (родовий відмінок)
            name_with_uppercase_lastname = f"{last_name_genitive.upper()} {first_name[0].upper()}."
            
            # Повне ім'я з формальним відмінком
            full_name = f"{rank} {name}"
        else:
            formal_name = name.upper()
            name_with_uppercase_lastname = name
            full_name = f"{rank} {name}"
        
        # Замінюємо плейсхолдери в параграфах
        for paragraph in doc.paragraphs:
            text = paragraph.text
            
            # Замінюємо всі плейсхолдери відповідними форматами
            text = text.replace("{name}", full_name)
            text = text.replace("{formal_name}", formal_name)
            text = text.replace("{short_name}", name_with_uppercase_lastname)
            
            # Інші заміни
            text = text.replace("{commander}", commander)
            text = text.replace("{rank}", rank)
            text = text.replace("{time_date}", time_date)
            text = text.replace("{reason}", reason)
            text = text.replace("{address}", address)
            text = text.replace("{report_date}", report_date)
            
            # Оновлюємо текст в параграфі
            paragraph.text = text

        # Зберігаємо новий файл
        output_filename = f"report_{name.replace(' ', '_').replace('.', '')}.docx"
        doc.save(output_filename)
        return output_filename
    except Exception as e:
        logging.error(f"Помилка при генерації рапорту: {e}")
        raise

# Запуск бота
async def main():
    try:
        # Перевіряємо наявність шаблону документа
        if not os.path.exists(template_path):
            logging.error(f"Шаблон документа '{template_path}' не знайдено!")
            print(f"Помилка: файл шаблону '{template_path}' не існує. Перевірте шлях до файлу.")
            return
        
        # Перевіряємо наявність файлів зі списками
        if not os.path.exists(names_path):
            logging.warning(f"Файл зі списком курсантів '{names_path}' не знайдено!")
            print(f"Попередження: файл '{names_path}' не існує. Буде використано порожній список.")
        
        if not os.path.exists(reasons_path):
            logging.warning(f"Файл зі списком причин '{reasons_path}' не знайдено!")
            print(f"Попередження: файл '{reasons_path}' не існує. Буде використано порожній список.")
            
        # Запускаємо бота
        logging.info("Бот запущено!")
        await dp.start_polling(bot)
    except Exception as e:
        logging.error(f"Помилка при запуску бота: {e}")

if __name__ == '__main__':
    logging.basicConfig(level=logging.INFO)
    asyncio.run(main())
