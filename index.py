
import pandas as pd
import smtplib
import imaplib
import email
import os
import time
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime, timedelta
from apscheduler.schedulers.blocking import BlockingScheduler
from email.header import decode_header
from email.header import decode_header
import re

# from datetime import datetime, timedelta
# from exchangelib import Message, Mailbox
# from exchangelib.items import SEND_ONLY_TO_ALL



# Конфигурация
SMTP_SERVER = "smtp.yandex.ru"
SMTP_PORT = 587
IMAP_SERVER = "imap.yandex.ru"
EMAIL = "zaharov.egor.2003@yandex.ru"  # Замените на реальный email
PASSWORD = "tqyxemaddulynkfc"  # Для Gmail используйте пароль приложения

# Проверка и создание файла tasks.xlsx
if not os.path.exists("tasks.xlsx"):
    df = pd.DataFrame(columns=["Задача", "Исполнитель", "Email", "Дедлайн", "Статус"])
    df.to_excel("tasks.xlsx", index=False)
    print("Создан новый файл tasks.xlsx")

def load_tasks():
    try:
        return pd.read_excel("tasks.xlsx")
    except FileNotFoundError:
        print("Файл tasks.xlsx не найден, создаю новый...")
        df = pd.DataFrame(columns=["Задача", "Исполнитель", "Email", "Дедлайн", "Статус"])
        df.to_excel("tasks.xlsx", index=False)
        return df

def send_email(to_email, task, assignee, deadline):
    try:
        msg = MIMEMultipart()
        msg['From'] = EMAIL
        msg['To'] = to_email
        msg['Subject'] = f"Напоминание: задача «{task}»"

        body = f"""Уважаемый(ая) {assignee},

                Напоминаем о задаче: «{task}».
                Срок выполнения: {deadline}.

                Ответьте на это письмо цифрой:
                123 — задача выполнена,
                321 — задача не выполнена."""
        msg.attach(MIMEText(body, 'plain'))

        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(EMAIL, PASSWORD)
            server.send_message(msg)
        print(f"Письмо отправлено на {to_email}")
    except Exception as e:
        print(f"Ошибка при отправке письма: {e}")

def check_deadlines():
    try:
        df = load_tasks()
        today = datetime.now().date()

        for index, row in df.iterrows():
            deadline = pd.to_datetime(row['Дедлайн']).date()
            if (deadline - today).days == 1:
                send_email(
                    to_email=row['Email'],
                    task=row['Задача'],
                    assignee=row['Исполнитель'],
                    deadline=row['Дедлайн']
                )
    except Exception as e:
        print(f"Ошибка в check_deadlines: {e}")


def decode_mime_header(header):
    """Декодирует MIME-заголовки"""
    if header is None:
        return ""
    decoded_parts = []
    for part, encoding in decode_header(header):
        if isinstance(part, bytes):
            decoded_parts.append(part.decode(encoding or 'utf-8', errors='replace'))
        else:
            decoded_parts.append(str(part))
    return "".join(decoded_parts)


def check_responses():
    try:
        print("\n Начинаем анализ ответных писем...")
        mail = imaplib.IMAP4_SSL(IMAP_SERVER)
        mail.login(EMAIL, PASSWORD)
        mail.select("inbox")

        status, messages = mail.search(None, "UNSEEN")
        if status != "OK" or not messages[0]:
            print("Нет новых непрочитанных писем")
            return

        message_ids = messages[0].split()
        print(f"Найдено {len(message_ids)} новых писем")

        for num in message_ids:
            try:
                print(f"\nАнализ письма ID: {num.decode()}")
                status, data = mail.fetch(num, "(RFC822)")
                if status != "OK":
                    continue

                msg = email.message_from_bytes(data[0][1])
                
                # Проверяем, что это ответ на наше письмо
                subject = decode_header(msg["Subject"])[0][0]
                if isinstance(subject, bytes):
                    subject = subject.decode('utf-8', errors='ignore')
                
                if not subject or "Re: Напоминание:" not in subject:
                    print("Пропускаем: не ответ на напоминание")
                    continue

                # Извлекаем email отправителя
                from_header = msg.get("From", "")
                from_email = re.search(r'[\w\.-]+@[\w\.-]+', from_header)
                if not from_email:
                    print("Не удалось извлечь email отправителя")
                    continue
                from_email = from_email.group(0)
                print(f"Отправитель: {from_email}")

                # Получаем текст письма
                body = ""
                if msg.is_multipart():
                    for part in msg.walk():
                        if part.get_content_type() == "text/plain":
                            body = part.get_payload(decode=True).decode('utf-8', errors='ignore')
                            break
                else:
                    body = msg.get_payload(decode=True).decode('utf-8', errors='ignore')

                # Нормализуем текст для анализа
                clean_body = re.sub(r'\s+', ' ', body).strip().lower()
                print(f"Текст письма: {clean_body[:200]}...")

                # Точный поиск статуса
                status = None
                if re.search('123', clean_body):
                    status = "Выполнено"
                elif re.search('321', clean_body):
                    status = "Не выполнено"
                else:
                    print("Не найдено цифр 123 или 321 в теле письма")
                    continue

                print(f"Определен статус: {status}")
                update_status(from_email, status)
                mail.store(num, "+FLAGS", "\\Seen")
                print("Письмо обработано")

            except Exception as e:
                print(f"Ошибка обработки письма: {str(e)}")
                continue

        mail.close()
        mail.logout()
    except Exception as e:
        print(f"Ошибка IMAP: {str(e)}")


def update_status(email, status):
    try:
        df = pd.read_excel("tasks.xlsx")
        # Ищем точное совпадение email (без учета регистра)
        mask = df["Email"].str.lower() == email.lower()
        if not any(mask):
            print(f" Email {email} не найден в tasks.xlsx")
            return
            
        df.loc[mask, "Статус"] = status
        df.to_excel("tasks.xlsx", index=False)
        print(f" Обновлен статус для {email}: {status}")
    except Exception as e:
        print(f" Ошибка при обновлении статуса: {e}")


def send_monthly_report():
    today = datetime.now().date()
    
    # Проверяем, что сегодня последний день месяца
    next_day = today + timedelta(days=1)
    # if next_day.month == today.month:
    if today != today:
        print("Сегодня не последний день месяца. Рассылка не требуется.")
        return
    
    print("⏳ Подготовка месячных отчетов...")
    
    # Загружаем данные
    try:
        df = load_tasks()
    except Exception as e:
        print(f"Ошибка загрузки задач: {e}")
        return
    
    # Группируем задачи по исполнителям
    grouped = df.groupby(['Email', 'Исполнитель'])
    
    # Настройки SMTP
    smtp_server = "smtp.yandex.ru"  # Для mail.ru (для других сервисов укажите свой)
    smtp_port = 587
    smtp_login = "zaharov.egor.2003@yandex.ru"  # Ваш email для отправки
    smtp_password = "tqyxemaddulynkfc"    # Пароль или app-пароль
    
    try:
        # Подключаемся к SMTP серверу
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(smtp_login, smtp_password)
            
            for (email, name), tasks in grouped:
                try:
                    # Проверяем валидность email
                    if not re.match(r'^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$', email):
                        print(f"Пропускаем невалидный email: {email}")
                        continue
                    
                    # Формируем список задач
                    task_list = []
                    for _, task in tasks.iterrows():
                        status = "Выполнено" if task['Статус'] == "Выполнено" else "Не выполнено"
                        deadline = pd.to_datetime(task['Дедлайн']).strftime('%d.%m.%Y')
                        task_list.append(f"\t{status} {task['Задача']} (до {deadline})")
                    
                    if not task_list:
                        continue
                        
                    # Создаем письмо
                    msg = MIMEMultipart()
                    msg['From'] = smtp_login
                    msg['To'] = email
                    msg['Subject'] = f"Ваши задачи на {today.strftime('%B %Y')}"
                    
                    # Текстовая версия
                    text = f"""Уважаемый(ая) {name},
                    
                            Ваши задачи на текущий месяц:
                    
                            """ + "\n".join(task_list) + """

                            С уважением,
                            Система учета задач
                            """
                    msg.attach(MIMEText(text, 'plain'))
                    
                    # Отправляем письмо
                    server.send_message(msg)
                    print(f"Отчет отправлен {name} <{email}>")
                    
                    # Пауза между отправкой писем
                    time.sleep(1)
                    
                except Exception as e:
                    print(f"Ошибка отправки отчета для {email}: {str(e)}")
                    
    except Exception as e:
        print(f"Ошибка подключения к SMTP серверу: {str(e)}")

def job():
    print("Запуск задачи...")
    check_deadlines()
    check_responses()
    send_monthly_report()

# if __name__ == "__main__":
#     print("Скрипт запущен")
#     try:
#         scheduler = BlockingScheduler()
#         scheduler.add_job(job, 'cron', hour=9)
#         print("Планировщик запущен. Ожидание задач...")
#         scheduler.start()
#     except KeyboardInterrupt:
#         print("Скрипт остановлен")
#     except Exception as e:
#         print(f"Ошибка: {e}")

job()