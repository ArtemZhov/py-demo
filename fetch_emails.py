#!/usr/bin/env python3
import os
import sys
import logging
import getpass
import keyring
import json
from exchangelib import Credentials, Account, DELEGATE, Configuration, ewsdatetime
from exchangelib.errors import TransportError, UnauthorizedError, EWSError
from exchangelib.attachments import FileAttachmentIO
from datetime import datetime
from bs4 import BeautifulSoup

logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")

# --- Настройки (замени username по необходимости или оставь пустым для ввода) ---
USERNAME = os.getenv("EMAIL_USER") or ""      # пример: name@mail.com
USE_AUTODISCOVER = False                       # False -> использовать EXPLICIT_SERVER
EXPLICIT_SERVER = "cas.rt.ru"     # в случае autodiscover=False
FETCH_COUNT = 10

OUTPUT_DIR = "fetched_emails_output" # Директория для сохранения файлов
ATTACHMENTS_SUBDIR = "attachments"   # Поддиректория для вложений внутри OUTPUT_DIR
SAVE_JSON = True                     # Сохранять ли данные в JSON файл
SAVE_CSV = True                      # Сохранять ли данные в CSV файл
SAVE_ATTACHMENTS = True              # Сохранять ли вложения
FETCH_ATTACHMENTS_CONTENT = True
retrieved_emails_data = []
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
    except EWSError as e:
        logging.error("Ошибка подключения к EWS: %s", e)
        logging.info("Если используете Exchange Online и базовая авторизация отключена — потребуется OAuth / Microsoft Graph.")
        sys.exit(2)
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
    try:
        logging.info("Получаем письма...")
        search_subject = "подобьемся"  # Упрощаем фильтр
        items = list(account.inbox.all().filter(subject__icontains=search_subject).order_by('-datetime_received')[:FETCH_COUNT])
        logging.info(f"Найдено писем: {len(items)}")
        for item in items:
            subj = item.subject or "(no subject)"
            sender = (item.sender.email_address if item.sender else "(no sender)")
            dt = item.datetime_received.isoformat()
            body = (item.text_body or item.body or "")[:200]
            # Выводим письмо в консоль
            print(f"Subject: {subj}")
            print(f"From: {sender}   Date: {dt}")
            print(body)
            print("-" * 72)
            # Добавляем данные письма в список
            retrieved_emails_data.append({
            "subject": subj,
            "from": sender,
            "date": dt,
            "body": body
        })
    except Exception as e:
        logging.error("Ошибка при чтении писем: %s", e)
        sys.exit(4)
    

# --- Сохранение в файлы ---
    if SAVE_JSON: # Проверяет, установлен ли флаг SAVE_JSON в True
                json_file_path = os.path.join(OUTPUT_DIR, "emails_data.json") # Формирует полный путь к JSON файлу
    with open(json_file_path, "w", encoding="utf-8") as f:
        # json.dump преобразует список объектов EmailData (предварительно преобразованных в словари через .to_dict()) в JSON формат
        json.dump(retrieved_emails_data, f, ensure_ascii=False, indent=4)
        logging.info(f"Данные писем сохранены в JSON файл: {json_file_path}") # Сообщение в консоль

if __name__ == "__main__":
    main()
