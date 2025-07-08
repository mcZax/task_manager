import pandas as pd
import smtplib
import imaplib
import email
import os
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime
from apscheduler.schedulers.blocking import BlockingScheduler
from email.header import decode_header
from email.header import decode_header
import re

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
                1 — задача выполнена,
                2 — задача не выполнена."""
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
        print("🔍 Проверяем новые ответы...")
        mail = imaplib.IMAP4_SSL(IMAP_SERVER)
        mail.login(EMAIL, PASSWORD)
        mail.select("inbox")

        # Ищем непрочитанные письма
        status, messages = mail.search(None, "UNSEEN")
        if status != "OK" or not messages[0]:
            print("Нет новых непрочитанных писем")
            return

        print(f"Найдено {len(messages[0].split())} новых писем")

        for num in messages[0].split():
            try:
                print(f"\nОбрабатываем письмо #{num.decode()}")
                status, data = mail.fetch(num, "(RFC822)")
                if status != "OK":
                    print(f"Ошибка получения письма #{num}")
                    continue

                msg = email.message_from_bytes(data[0][1])

                subject = decode_mime_header(msg.get("Subject"))
                print(f"Декодированная тема: {subject}")  # Для отладки
            
                # Проверяем тему через регулярное выражение
                if not re.search(r"Re:\s*Напоминание:", subject, re.IGNORECASE):
                    continue
                
                # Получаем отправителя
                from_header = msg.get("From", "")
                from_email = ""
                if "<" in from_header and ">" in from_header:
                    from_email = from_header.split("<")[1].split(">")[0]
                else:
                    from_email = from_header
                print(f"От: {from_email}")

                # Получаем тему
                from_header = msg.get("From", "")
                from_email = re.search(r'<(.+?)>', from_header) or re.search(r'(\S+@\S+)', from_header)
                from_email = from_email.group(1) if from_email else ""

                # Пропускаем если это не ответ на наше письмо
                if not subject.startswith("Re: Напоминание:"):
                    print("Пропускаем - не наш ответ")
                    continue

                # Извлекаем тело письма
                body = ""
                if msg.is_multipart():
                    for part in msg.walk():
                        content_type = part.get_content_type()
                        content_disposition = str(part.get("Content-Disposition"))
                        if content_type == "text/plain" and "attachment" not in content_disposition:
                            body = part.get_payload(decode=True).decode()
                            break
                else:
                    body = msg.get_payload(decode=True).decode()

                print(f"Тело письма:\n{body[:200]}...")  # Логируем начало тела

                clean_body = re.sub(r'\s+', '', body)
                
                # Ищем точные совпадения цифр
                if re.search(r'(^|[^0-9])1([^0-9]|$)', clean_body):
                    print(f"Найдено подтверждение выполнения от {from_email}")
                    update_status(from_email, "Выполнено")
                elif re.search(r'(^|[^0-9])2([^0-9]|$)', clean_body):
                    print(f"Найдено подтверждение НЕ выполнения от {from_email}")
                    update_status(from_email, "Не выполнено")
                else:
                    print(f"Не найдено цифр 1 или 2 в письме от {from_email}")

                # Помечаем как прочитанное
                mail.store(num, "+FLAGS", "\\Seen")
                print("Письмо обработано и помечено как прочитанное")

            except Exception as e:
                print(f"⚠️ Ошибка при обработке письма #{num}: {e}")
                continue

        mail.close()
        mail.logout()
        print("Проверка ответов завершена")
    except Exception as e:
        print(f"❌ Критическая ошибка в check_responses: {e}")


def update_status(email, status):
    try:
        df = pd.read_excel("tasks.xlsx")
        # Ищем точное совпадение email (без учета регистра)
        mask = df["Email"].str.lower() == email.lower()
        if not any(mask):
            print(f"⚠️ Email {email} не найден в tasks.xlsx")
            return
            
        df.loc[mask, "Статус"] = status
        df.to_excel("tasks.xlsx", index=False)
        print(f"✅ Обновлен статус для {email}: {status}")
    except Exception as e:
        print(f"❌ Ошибка при обновлении статуса: {e}")

def job():
    print("Запуск задачи...")
    check_deadlines()
    check_responses()

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