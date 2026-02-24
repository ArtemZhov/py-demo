#!/usr/bin/env python3
import os
import sys
import logging
import getpass
import keyring
import json
from exchangelib import Credentials, Account, DELEGATE, Configuration, FileAttachment
from datetime import datetime, timezone
import email
import re

logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")

# --- Настройки путей --- 
# Определяем директорию, где находится этот скрипт
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# --- Настройки (замени username по необходимости или оставь пустым для ввода) ---
USERNAME = os.getenv("EMAIL_USER") or ""      # пример: name@mail.com
USE_AUTODISCOVER = False                       # False -> использовать EXPLICIT_SERVER
EXPLICIT_SERVER = "cas.rt.ru"     # в случае autodiscover=False

# Список тематик для поиска (можно добавить любое количество)
SEARCH_TOPICS = [
    "важная информация по проекту шашлыки",
    "Вопрос по проекту шашлыков",
    "Протокол встречи по проекту Кролики",
]

# Увеличиваем лимит или убираем его, чтобы получать больше писем.
# Если нужно абсолютно все письма, можно убрать срез [:FETCH_COUNT] в коде ниже.
FETCH_COUNT = 1000 

# Формируем абсолютный путь к папке с результатами
OUTPUT_DIR = os.path.join(SCRIPT_DIR, "fetched_emails_output")
ATTACHMENTS_SUBDIR = "attachments"   # Поддиректория для вложений внутри OUTPUT_DIR
SAVE_JSON = True                     # Сохранять ли данные в JSON файл
SAVE_CSV = True                      # Сохранять ли данные в CSV файл
SAVE_ATTACHMENTS = True              # Сохранять ли вложения
FETCH_ATTACHMENTS_CONTENT = True

def sanitize_filename(filename):
    """Удаляет или заменяет недопустимые символы в имени файла."""
    return re.sub(r'[\\/*?:"<>|]', "", filename)

def get_password(username: str) -> str:
    # 1) keyring
    pw = keyring.get_password("my_email_app", username)
    if pw:
        logging.info("Password loaded from keyring.")
        return pw
    # 2) env
    pw = os.getenv("EMAIL_PASS")
    if pw:
        logging.info("Password loaded from environment variable EMAIL_PASS.")
        return pw
    # 3) interactive (не печатается)
    while True:
        pw = getpass.getpass(f"Password for {username}: ")
        if pw:
            return pw
        print("Пароль не может быть пустым. Повторите ввод.")

def get_thread_id(msg):
    """
    Определяет ID переписки на основе заголовков.
    Приоритет: References -> In-Reply-To -> Message-ID.
    """
    refs = msg.get("References")
    if refs:
        # Берем самый старый ID из цепочки References как корень
        return refs.split()[-1] if refs.split() else None

    in_reply_to = msg.get("In-Reply-To")
    if in_reply_to:
        return in_reply_to.strip()

    # Если это начало новой цепочки
    msg_id = msg.get("Message-ID")
    return msg_id.strip() if msg_id else None

def parse_date(msg):
    """Вспомогательная функция для парсинга даты письма."""
    date_str = msg.get("Date")
    if not date_str:
        return datetime.min
    try:
        return email.utils.parsedate_to_datetime(date_str)
    except:
        return datetime.min

def build_email_tree(messages):
    """
    Строит древовидную структуру писем на основе Message-ID и In-Reply-To.
    """
    # Создаем словарь для быстрого доступа к сообщениям по ID
    messages_map = {m['id']: m for m in messages}
    roots = []

    # Предварительная инициализация списка children для всех сообщений,
    # чтобы избежать KeyError при обращении к родителю, который еще не обработан.
    for msg in messages:
        msg['children'] = []

    for msg in messages:
        parent_id = msg.get('in_reply_to')
        
        # Если есть ID родителя и он найден среди загруженных писем, добавляем текущее сообщение в children родителя
        if parent_id and parent_id in messages_map:
            messages_map[parent_id]['children'].append(msg)
        else:
            # Иначе считаем сообщение корневым (началом ветки)
            roots.append(msg)
    
    # Сортируем детей по дате внутри каждого узла для правильного порядка
    for msg in messages:
        msg['children'].sort(key=lambda x: x['date'])

    # Сортируем корневые сообщения по дате
    roots.sort(key=lambda x: x['date'])
    
    return roots

def clean_email_text(text: str) -> str:
    """
    Очищает текст письма от спецсимволов, например литеральных \r\n,
    и убирает лишние пустые строки.
    """
    if not text:
        return ""
    
    # Заменяем литеральные последовательности \r\n на реальные переносы строк
    text = text.replace('\\r\\n', '\n')
    text = text.replace('\\r', '\n')
    
    # Разбиваем на строки, убираем пробелы по краям и фильтруем пустые
    lines = [line.strip() for line in text.split('\n')]
    non_empty_lines = [line for line in lines if line]
    
    # Собираем обратно, разделяя абзацы одним переносом строки
    return '\n'.join(non_empty_lines)

def serialize_for_json(obj):
    """
    Рекурсивно преобразует объекты в JSON-совместимые типы.
    """
    if isinstance(obj, datetime):
        return obj.isoformat()
    elif isinstance(obj, bytes):
        return "<binary data>"
    elif isinstance(obj, list):
        return [serialize_for_json(item) for item in obj]
    elif isinstance(obj, dict):
        return {key: serialize_for_json(val) for key, val in obj.items()}
    elif hasattr(obj, '__dict__'):
        # Для объектов exchangelib
        return str(obj)
    else:
        return obj

def main():
    username = USERNAME.strip() or input("Email (e.g. name@mail.com): ").strip()
    if not username:
        print("Email не указан. Выход.")
        sys.exit(1)

    password = get_password(username)

    creds = Credentials(username=username, password=password)

    try:
        if USE_AUTODISCOVER:
            logging.info("Connecting with autodiscover...")
            account = Account(primary_smtp_address=username, credentials=creds,
                              autodiscover=True, access_type=DELEGATE)
        else:
            logging.info(f"Connecting to explicit server {EXPLICIT_SERVER} ...")
            config = Configuration(server=EXPLICIT_SERVER, credentials=creds)
            account = Account(primary_smtp_address=username, config=config,
                              autodiscover=False, access_type=DELEGATE)
    except Exception as e:
        logging.error("Неожиданная ошибка при подключении: %s", e)
        sys.exit(3)

    # --- Подготовка к сохранению ---
    os.makedirs(OUTPUT_DIR, exist_ok=True) # Создает корневую папку для выходных данных
    if SAVE_ATTACHMENTS: # Проверяет, нужно ли сохранять вложения
        attachments_full_path = os.path.join(OUTPUT_DIR, ATTACHMENTS_SUBDIR) # Формирует путь для вложений
        os.makedirs(attachments_full_path, exist_ok=True) # Создает папку для вложений
    else:
        attachments_full_path = None # Если не сохраняем, путь - None

    # Инициализируем структуру для хранения данных по всем тематикам
    # Структура: { "topics": [ { "topic": "тема", "emails": [...], "threads": [...] }, ... ] }
    final_data = {
        "topics": []
    }

    try:
        # --- Поиск писем по каждой тематике ---
        for search_topic in SEARCH_TOPICS:
            logging.info(f"Поиск писем по теме: {search_topic}")
            
            # Получаем письма для текущей темы
            items = list(account.inbox.all().filter(subject__icontains=search_topic).order_by('-datetime_received'))
            logging.info(f"Найдено писем по теме '{search_topic}': {len(items)}")
            
            # Пропускаем тему, если писем не найдено
            if not items:
                continue
            
            # Группируем письма по ID переписки для текущей темы
            threads = {}
            topic_emails = []  # Плоский список писем для текущей темы
            
            for item in items:
                try:
                    # Проверяем наличие MIME контента, чтобы избежать падения
                    if not item.mime_content:
                        logging.warning(f"Письмо без темы или контента пропущено (ID: {item.id}).")
                        continue

                    subj = item.subject or "(no subject)"
                    sender = (item.sender.email_address if item.sender else "(no sender)")
                    dt = item.datetime_received.isoformat()
                    body = item.text_body or item.body or ""

                    # --- Обработка вложений ---
                    attachments_info = []
                    if SAVE_ATTACHMENTS and item.attachments and attachments_full_path:
                        for attachment in item.attachments:
                            if isinstance(attachment, FileAttachment):
                                try:
                                    # Генерируем безопасное имя файла
                                    safe_name = sanitize_filename(attachment.name)
                                    # Добавляем префикс на основе ID письма, чтобы избежать конфликтов имен
                                    prefix = item.id.replace('-', '_')[:10]
                                    file_name = f"{prefix}_{safe_name}"
                                    file_path = os.path.join(attachments_full_path, file_name)
                                    
                                    # Сохраняем файл
                                    with open(file_path, 'wb') as f:
                                        f.write(attachment.content)
                                    
                                    attachments_info.append({
                                        "name": attachment.name,
                                        "saved_as": file_name,
                                        "path": file_path
                                    })
                                    logging.info(f"Сохранено вложение: {file_name}")
                                except Exception as att_err:
                                    logging.error(f"Ошибка сохранения вложения {attachment.name}: {att_err}")

                    # Выводим письмо в консоль
                    print(f"Subject: {subj}")
                    print(f"From: {sender}   Date: {dt}")
                    if attachments_info:
                        print(f"Attachments: {len(attachments_info)}")
                    print(body)
                    print("-" * 72)

                    # Получаем raw MIME контент для парсинга заголовков
                    raw_email = item.mime_content
                    msg = email.message_from_bytes(raw_email)
                    
                    msg_id = msg.get("Message-ID")
                    in_reply_to = msg.get("In-Reply-To")

                    current_msg_id = msg_id.strip() if msg_id else None
                    parent_id = in_reply_to.strip() if in_reply_to else None
                    thread_id = get_thread_id(msg)

                    # Обновляем словарь данных письма, добавляя информацию о вложениях и теме
                    email_data = {
                        "id": current_msg_id,
                        "in_reply_to": parent_id,
                        "subject": subj,
                        "from": sender,
                        "body": body,
                        "date": item.datetime_received,
                        "attachments": attachments_info,
                        "topic": search_topic  # Добавлено: связь с темой
                    }

                    # Добавляем письмо в плоский список
                    topic_emails.append(email_data)

                    # Добавляем письмо в словарь threads
                    if thread_id:
                        if thread_id not in threads:
                            threads[thread_id] = []
                        threads[thread_id].append(email_data)

                except Exception as item_error:
                    logging.error(f"Ошибка при обработке письма {item.subject}: {item_error}")
                    continue # Продолжаем со следующим письмом

            # --- Построение дерева для текущей темы ---
            topic_tree = []
            
            if threads:
                logging.info(f"Найдено переписок по теме '{search_topic}': {len(threads)}")
                
                for thread_messages in threads.values():
                    # Конвертируем объекты datetime в строки ISO format для JSON
                    for msg in thread_messages:
                        if isinstance(msg['date'], datetime):
                            msg['date'] = msg['date'].isoformat()
                    
                    # Строим дерево для текущей переписки и добавляем корни в общий список
                    thread_roots = build_email_tree(thread_messages)
                    topic_tree.extend(thread_roots)
            else:
                logging.warning(f"Переписки не найдены по теме '{search_topic}'.")
            # Сортируем корневые сообщения по дате
            topic_tree.sort(key=lambda x: x['date'])

            # Добавляем данные темы в общую структуру
            final_data["topics"].append({
                "topic": search_topic,
                "emails": topic_emails,  # Плоский список
                "threads": topic_tree    # Древовидная структура
            })

    except Exception as e:
        logging.error("Критическая ошибка при чтении писем: %s", e)

    # --- Сохранение в файлы ---
    # Вынесено из try-except, чтобы выполнялось всегда
    if SAVE_JSON:
        json_file_path = os.path.join(OUTPUT_DIR, "emails_data.json")
        logging.info(f"Попытка сохранения в файл: {json_file_path}")
        
        try:
            with open(json_file_path, "w", encoding="utf-8") as f:
                json.dump(final_data, f, ensure_ascii=False, indent=4)
            logging.info(f"Данные писем успешно сохранены в JSON файл: {json_file_path}")
        except (IOError, PermissionError) as e:
            logging.error(f"Ошибка доступа к файлу {json_file_path}: {e}")
        except TypeError as e:
            logging.error(f"Ошибка сериализации данных в JSON: {e}")

    return final_data
    
if __name__ == "__main__":
    main()