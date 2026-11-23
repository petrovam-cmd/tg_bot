import pandas as pd
import logging
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes
from collections import defaultdict, Counter
import re
from datetime import datetime
import os
from dotenv import load_dotenv

# Загрузка переменных окружения
load_dotenv()

# Настройка логирования
logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    level=logging.INFO
)

# Ключевые слова для категоризации проблем
PROBLEM_KEYWORDS = {
    'отопление': ['отоплен', 'батаре', 'тепло', 'холодно', 'обогрев', 'отопительн'],
    'уборка': ['уборк', 'мусор', 'гряз', 'чисто', 'подмет', 'убирать'],
    'бытовые аварии': ['авария', 'протеч', 'затоп', 'поломка', 'прорвало', 'засор'],
    'дороги': ['яма', 'дорог', 'асфальт', 'выбоин', 'покрытие', 'тротуар'],
    'мусорные баки': ['бак', 'контейнер', 'свалка', 'мусорк', 'вывоз мусора'],
    'жалобы': ['жалоб', 'плох', 'ремонт', 'неисправн', 'проблем']
}

# Регулярные выражения для поиска адресов
ADDRESS_PATTERNS = [
    r'(ул\.?|улица|просп\.?|проспект|пер\.?|переулок|б-р|бульвар)\s+[А-Яа-я\w\s\-]+',  # улицы
    r'д\.?\s*\d+',  # дома
    r'кв\.?\s*\d+',  # квартиры
    r'дом\s*\d+',
]


class ProblemTrackerBot:
    def __init__(self, token: str, excel_file: str = "Форма.xlsx", top_excel_file: str = "топ_адресов.xlsx"):
        self.token = token
        self.excel_file = excel_file
        self.top_excel_file = top_excel_file
        self.message_counter = Counter()
        self.message_reactions = defaultdict(dict)  # {message_text: {reaction_type: count}}
        self.problem_messages = defaultdict(list)

        # Загружаем существующие данные
        self.load_existing_data()

    def load_existing_data(self):
        """Загружает существующие данные из Excel файлов"""
        try:
            # Загрузка основного файла
            if os.path.exists(self.excel_file):
                self.df = pd.read_excel(self.excel_file)
                # Обновляем счетчик сообщений и реакций
                for _, row in self.df.iterrows():
                    if pd.notna(row['Текст сообщения']):
                        message_text = row['Текст сообщения']
                        self.message_counter[message_text] = row['Количество повторений']

                        # Загружаем реакции если они есть
                        if 'Реакции' in row and pd.notna(row['Реакции']):
                            try:
                                # Безопасная загрузка словаря реакций из строки
                                reactions_str = row['Реакции'].replace("'", '"')
                                import json
                                reactions = json.loads(reactions_str)
                                self.message_reactions[message_text] = reactions
                            except:
                                self.message_reactions[message_text] = {}
            else:
                # Создаем новый DataFrame с колонкой для реакций
                self.df = pd.DataFrame(columns=[
                    'ID', 'Номер телефона/имя пользователя', 'ФИО (если есть)',
                    'Текст сообщения', 'Открытие (да/нет)', 'Категория',
                    'Адрес', 'Дата', 'Количество повторений', 'Реакции', 'Общее количество реакций'
                ])

            # Загрузка файла с топом адресов (если существует)
            if os.path.exists(self.top_excel_file):
                self.top_df = pd.read_excel(self.top_excel_file)
            else:
                self.top_df = pd.DataFrame(columns=[
                    'Текст сообщения', 'Категория', 'Адрес',
                    'Количество повторений', 'Общее количество реакций',
                    'Реакции', 'Дата последнего упоминания'
                ])

        except Exception as e:
            logging.error(f"Ошибка загрузки Excel: {e}")
            self.df = pd.DataFrame(columns=[
                'ID', 'Номер телефона/имя пользователя', 'ФИО (если есть)',
                'Текст сообщения', 'Открытие (да/нет)', 'Категория',
                'Адрес', 'Дата', 'Количество повторений', 'Реакции', 'Общее количество реакций'
            ])
            self.top_df = pd.DataFrame(columns=[
                'Текст сообщения', 'Категория', 'Адрес',
                'Количество повторений', 'Общее количество реакций',
                'Реакции', 'Дата последнего упоминания'
            ])

    def save_to_excel(self):
        """Сохраняет данные в основной Excel файл"""
        try:
            # Обновляем количество повторений и реакций перед сохранением
            for idx, row in self.df.iterrows():
                message_text = row['Текст сообщения']
                if pd.notna(message_text):
                    self.df.at[idx, 'Количество повторений'] = self.message_counter[message_text]
                    self.df.at[idx, 'Реакции'] = str(self.message_reactions.get(message_text, {}))
                    self.df.at[idx, 'Общее количество реакций'] = self.calculate_total_reactions(
                        self.message_reactions.get(message_text, {})
                    )

            self.df.to_excel(self.excel_file, index=False)
            logging.info("Данные сохранены в основной Excel файл")
        except Exception as e:
            logging.error(f"Ошибка сохранения в основной Excel: {e}")

    def save_top_messages_to_excel(self):
        """Сохраняет топ сообщений в отдельный Excel файл"""
        try:
            # Получаем топ сообщений (5+ повторений)
            top_messages = self.get_top_messages(min_repeats=5)

            # Создаем новый DataFrame для топа
            top_data = []
            for msg in top_messages:
                # Находим последнюю дату упоминания этого сообщения
                last_mention_date = self.df[
                    self.df['Текст сообщения'] == msg['text']
                    ]['Дата'].max()

                top_data.append({
                    'Текст сообщения': msg['text'],
                    'Категория': self.categorize_message(msg['text']),
                    'Адрес': self.extract_address(msg['text']),
                    'Количество повторений': msg['repeats'],
                    'Общее количество реакций': msg['total_reactions'],
                    'Реакции': str(msg['reactions']),
                    'Дата последнего упоминания': last_mention_date if pd.notna(
                        last_mention_date) else datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                })

            # Сортируем по количеству повторений (по убыванию)
            top_data.sort(key=lambda x: x['Количество повторений'], reverse=True)

            self.top_df = pd.DataFrame(top_data)
            self.top_df.to_excel(self.top_excel_file, index=False)
            logging.info(f"Топ сообщений сохранен в {self.top_excel_file}")

        except Exception as e:
            logging.error(f"Ошибка сохранения топа в Excel: {e}")

    def categorize_message(self, text: str) -> str:
        """Определяет категорию сообщения"""
        text_lower = text.lower()
        for category, keywords in PROBLEM_KEYWORDS.items():
            for keyword in keywords:
                if keyword in text_lower:
                    return category
        return 'другое'

    def extract_address(self, text: str) -> str:
        """Извлекает адрес из текста сообщения"""
        addresses = []
        for pattern in ADDRESS_PATTERNS:
            matches = re.findall(pattern, text, re.IGNORECASE)
            addresses.extend(matches)

        # Объединяем найденные части адреса
        if addresses:
            return ', '.join(addresses)
        return ''

    def calculate_total_reactions(self, reactions_dict: dict) -> int:
        """Вычисляет общее количество реакций"""
        return sum(reactions_dict.values()) if reactions_dict else 0

    def process_message(self, message_text: str, user_info: dict) -> dict:
        """Обрабатывает сообщение и возвращает данные для записи"""
        # Увеличиваем счетчик для этого сообщения
        self.message_counter[message_text] += 1

        # Определяем категорию
        category = self.categorize_message(message_text)

        # Извлекаем адрес
        address = self.extract_address(message_text)

        # Получаем текущие реакции для этого сообщения
        current_reactions = self.message_reactions.get(message_text, {})
        total_reactions = self.calculate_total_reactions(current_reactions)

        return {
            'ID': len(self.df) + 1,
            'Номер телефона/имя пользователя': user_info.get('username', ''),
            'ФИО (если есть)': user_info.get('full_name', ''),
            'Текст сообщения': message_text,
            'Открытие (да/нет)': 'нет',
            'Категория': category,
            'Адрес': address,
            'Дата': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'Количество повторений': self.message_counter[message_text],
            'Реакции': str(current_reactions),
            'Общее количество реакций': total_reactions
        }

    def add_reaction(self, message_text: str, reaction_emoji: str):
        """Добавляет реакцию к сообщению"""
        if message_text not in self.message_reactions:
            self.message_reactions[message_text] = {}

        self.message_reactions[message_text][reaction_emoji] = (
                self.message_reactions[message_text].get(reaction_emoji, 0) + 1
        )

        # Обновляем данные в DataFrame
        if message_text in self.df['Текст сообщения'].values:
            idx = self.df[self.df['Текст сообщения'] == message_text].index[0]
            self.df.at[idx, 'Реакции'] = str(self.message_reactions[message_text])
            self.df.at[idx, 'Общее количество реакций'] = self.calculate_total_reactions(
                self.message_reactions[message_text]
            )

    async def handle_message_with_reaction(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Обработчик сообщений с реакциями (эмулируем сбор реакций через ответы)"""
        try:
            user = update.message.from_user
            message_text = update.message.text

            # Пропускаем команды
            if message_text.startswith('/'):
                return

            # Проверяем, является ли это сообщение реакцией (ответом на другое сообщение)
            if update.message.reply_to_message:
                original_message = update.message.reply_to_message.text
                reaction_text = message_text.lower()

                # Определяем тип реакции по ключевым словам
                reaction_emojis = {
                    'like': '👍',
                    'dislike': '👎',
                    'heart': '❤️',
                    'fire': '🔥',
                    'ok': '👌',
                    'plus': '➕',
                    'minus': '➖'
                }

                # Ищем ключевые слова в тексте реакции
                detected_emoji = None
                for reaction_word, emoji in reaction_emojis.items():
                    if reaction_word in reaction_text:
                        detected_emoji = emoji
                        break

                if detected_emoji:
                    # Добавляем реакцию к оригинальному сообщению
                    self.add_reaction(original_message, detected_emoji)
                    self.save_to_excel()
                    self.save_top_messages_to_excel()  # Обновляем топ

                    await update.message.reply_text(
                        f"Реакция {detected_emoji} добавлена к сообщению!"
                    )
                    return

            # Обычная обработка сообщения
            user_info = {
                'username': user.username or f"user_{user.id}",
                'full_name': f"{user.first_name or ''} {user.last_name or ''}".strip()
            }

            message_data = self.process_message(message_text, user_info)

            # Обновляем или добавляем запись
            if message_text in self.df['Текст сообщения'].values:
                idx = self.df[self.df['Текст сообщения'] == message_text].index[0]
                for col, value in message_data.items():
                    if col != 'ID':  # Не обновляем ID
                        self.df.at[idx, col] = value
            else:
                new_row = pd.DataFrame([message_data])
                self.df = pd.concat([self.df, new_row], ignore_index=True)

            self.save_to_excel()
            self.save_top_messages_to_excel()  # Обновляем топ после нового сообщения
            logging.info(f"Обработано сообщение от {user_info['username']}: {message_text[:50]}...")

        except Exception as e:
            logging.error(f"Ошибка обработки сообщения: {e}")

    def get_top_messages(self, min_repeats: int = 5):
        """Возвращает сообщения, которые повторяются min_repeats раз и более, отсортированные по популярности"""
        # Фильтруем сообщения с достаточным количеством повторений
        filtered_messages = []

        for message_text, count in self.message_counter.items():
            if count >= min_repeats:
                reactions = self.message_reactions.get(message_text, {})
                total_reactions = self.calculate_total_reactions(reactions)

                filtered_messages.append({
                    'text': message_text,
                    'repeats': count,
                    'reactions': reactions,
                    'total_reactions': total_reactions,
                    'score': count + total_reactions  # Комбинированный score для сортировки
                })

        # Сортируем по количеству повторений (по убыванию), затем по реакциям (по убыванию)
        filtered_messages.sort(key=lambda x: (-x['repeats'], -x['total_reactions']))

        return filtered_messages

    async def start(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Обработчик команды /start"""
        await update.message.reply_text(
            "🤖 Бот для отслеживания проблемных сообщений активирован!\n\n"
            "Доступные команды:\n"
            "/stats - статистика по сообщениям\n"
            "/problems - показать частые проблемы\n"
            "/export - экспорт в Excel\n"
            "/top5 - топ-5 самых частых проблем\n"
            "/top_repeated - топ сообщений (5+ повторений)\n"
            "/export_top - экспорт топа сообщений\n\n"
            "💡 Чтобы добавить реакцию к сообщению, ответьте на него одним из слов:\n"
            "like, dislike, heart, fire, ok, plus, minus"
        )

    async def show_stats(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Показывает статистику по сообщениям"""
        try:
            total_messages = len(self.df)
            if total_messages == 0:
                await update.message.reply_text("📊 Пока нет обработанных сообщений")
                return

            # Статистика по категориям
            category_stats = self.df['Категория'].value_counts()

            response = "📊 Статистика сообщений:\n\n"
            response += f"Всего сообщений: {total_messages}\n\n"
            response += "По категориям:\n"

            for category, count in category_stats.items():
                response += f"  {category}: {count}\n"

            # Топ повторяющихся сообщений
            top_repeated = self.df.nlargest(3, 'Количество повторений')
            if not top_repeated.empty:
                response += "\n🔝 Самые частые сообщения:\n"
                for _, row in top_repeated.iterrows():
                    reactions_info = f" (реакций: {row['Общее количество реакций']})" if row[
                                                                                             'Общее количество реакций'] > 0 else ""
                    response += f"  {row['Текст сообщения'][:50]}... ({row['Количество повторений']} раз{reactions_info})\n"

            await update.message.reply_text(response)

        except Exception as e:
            await update.message.reply_text(f"❌ Ошибка при получении статистики: {e}")

    async def show_problems(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Показывает сообщения с проблемами"""
        try:
            # Фильтруем сообщения с проблемами (исключаем 'другое')
            problem_df = self.df[self.df['Категория'] != 'другое']

            if problem_df.empty:
                await update.message.reply_text("⚠️ Проблемные сообщения не найдены")
                return

            response = "🚨 Сообщения с проблемами:\n\n"

            for category in PROBLEM_KEYWORDS.keys():
                category_messages = problem_df[problem_df['Категория'] == category]
                if not category_messages.empty:
                    response += f"🔹 {category.upper()}:\n"

                    # Берем до 3 сообщений для каждой категории
                    for _, row in category_messages.head(3).iterrows():
                        address_info = f" (адрес: {row['Адрес']})" if row['Адрес'] else ""
                        reactions_info = f" [реакций: {row['Общее количество реакций']}]" if row[
                                                                                                 'Общее количество реакций'] > 0 else ""
                        response += f"   • {row['Текст сообщения'][:60]}...{address_info}{reactions_info}\n"
                    response += "\n"

            await update.message.reply_text(response[:4000])

        except Exception as e:
            await update.message.reply_text(f"❌ Ошибка при получении проблем: {e}")

    async def export_data(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Экспортирует основные данные в Excel и отправляет файл"""
        try:
            self.save_to_excel()

            with open(self.excel_file, 'rb') as file:
                await update.message.reply_document(
                    document=file,
                    filename=self.excel_file,
                    caption="📊 Экспорт основных данных в Excel"
                )

        except Exception as e:
            await update.message.reply_text(f"❌ Ошибка при экспорте: {e}")

    async def export_top_data(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Экспортирует топ сообщений в Excel и отправляет файл"""
        try:
            self.save_top_messages_to_excel()

            with open(self.top_excel_file, 'rb') as file:
                await update.message.reply_document(
                    document=file,
                    filename=self.top_excel_file,
                    caption="🏆 Экспорт топа сообщений в Excel"
                )

        except Exception as e:
            await update.message.reply_text(f"❌ Ошибка при экспорте топа: {e}")

    async def show_top5(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Показывает топ-5 самых частых проблемных сообщений"""
        try:
            # Группируем по тексту сообщения и берем самые частые
            top_messages = self.df.nlargest(5, 'Количество повторений')

            if top_messages.empty:
                await update.message.reply_text("📝 Пока нет частых сообщений")
                return

            response = "🏆 Топ-5 самых частых проблемных сообщений:\n\n"

            for i, (_, row) in enumerate(top_messages.iterrows(), 1):
                address_info = f"\n   📍 Адрес: {row['Адрес']}" if row['Адрес'] else ""
                reactions_info = f"\n   👍 Реакций: {row['Общее количество реакций']}" if row[
                                                                                             'Общее количество реакций'] > 0 else ""
                response += f"{i}. {row['Текст сообщения'][:80]}...\n"
                response += f"   🏷 Категория: {row['Категория']}{address_info}{reactions_info}\n"
                response += f"   🔢 Повторений: {row['Количество повторений']}\n\n"

            await update.message.reply_text(response)

        except Exception as e:
            await update.message.reply_text(f"❌ Ошибка при получении топа: {e}")

    async def show_top_repeated(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Показывает топ сообщений, которые повторяются 5+ раз, отсортированные по популярности"""
        try:
            top_messages = self.get_top_messages(min_repeats=5)

            if not top_messages:
                await update.message.reply_text("📝 Нет сообщений, повторяющихся 5 и более раз")
                return

            response = "🏆 Топ сообщений (5+ повторений, отсортированы по популярности):\n\n"

            for i, msg in enumerate(top_messages[:10], 1):  # Показываем топ-10
                reactions_info = f" [Реакций: {msg['total_reactions']}]" if msg['total_reactions'] > 0 else ""
                response += f"{i}. {msg['text'][:80]}...\n"
                response += f"   🔢 Повторений: {msg['repeats']}{reactions_info}\n"
                response += f"   ⭐ Комбинированный score: {msg['score']}\n\n"

            await update.message.reply_text(response[:4000])

        except Exception as e:
            await update.message.reply_text(f"❌ Ошибка при получении топа: {e}")

    def run(self):
        """Запускает бота"""
        application = Application.builder().token(self.token).build()

        # Добавляем обработчики
        application.add_handler(CommandHandler("start", self.start))
        application.add_handler(CommandHandler("stats", self.show_stats))
        application.add_handler(CommandHandler("problems", self.show_problems))
        application.add_handler(CommandHandler("export", self.export_data))
        application.add_handler(CommandHandler("export_top", self.export_top_data))
        application.add_handler(CommandHandler("top5", self.show_top5))
        application.add_handler(CommandHandler("top_repeated", self.show_top_repeated))

        # Обработчик для всех сообщений (включая реакции через ответы)
        application.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, self.handle_message_with_reaction))

        # Запускаем бота
        print("Бот запущен...")
        application.run_polling()


# Получение токена из переменных окружения
BOT_TOKEN = os.getenv('BOT_TOKEN')

if not BOT_TOKEN:
    raise ValueError("Токен бота не найден в переменных окружения. Проверьте файл .env")

# Запуск бота
if __name__ == "__main__":
    bot = ProblemTrackerBot(BOT_TOKEN)
    bot.run()