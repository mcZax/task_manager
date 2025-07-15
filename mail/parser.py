import pandas as pd
import imaplib
import email
from email.header import decode_header
from email.header import decode_header
import re
from config import IMAP_SERVER, EMAIL, PASSWORD, EXCEL_FILE, DATE_LOG
from datetime import datetime
from database.excel_handler import load_tasks, init_log
from mail.sender import send_email


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


def log_received_task(task):
    try:
        log_df = init_log()
        
        # Проверяем, существует ли уже такая задача
        task_exists = log_df["Задача"].eq(task).any()
        
        if task_exists:
            # Обновляем время для существующей задачи
            log_df.loc[log_df["Задача"] == task, 
                      "Дата получения ответа"] = datetime.now().strftime("%Y-%m-%d %H:%M")
        else:
            # Добавляем новую запись
            new_entry = {
                "Задача": task,
                "Дата и время напоминания": None,
                "Дата получения ответа": datetime.now().strftime("%Y-%m-%d %H:%M"),
            }
            log_df = pd.concat([log_df, pd.DataFrame([new_entry])], ignore_index=True)
        
        # Сохраняем изменения
        log_df.to_excel(DATE_LOG, index=False)
        
    except Exception as e:
        print(f"Ошибка при логировании задачи: {e}")
        raise


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
                # print(f"Текст письма: {clean_body}...")

                print(clean_body)

                task = re.search(r'«(.+?)»', clean_body).group(1)
                print(task)
                # Точный поиск статуса
                status = None
                if re.search('123', clean_body[:30]):
                    status = "Выполнено"
                elif re.search('321', clean_body[:30]):
                    status = "Не выполнено"
                else:
                    print("Не найдено цифр 123 или 321 в теле письма")
                    continue

                print(f"Определен статус: {status}")
                update_status(task, status)

                log_received_task(task)
                print("Дата и время получения записано в лог")
                mail.store(num, "+FLAGS", "\\Seen")
                print("Письмо обработано")

            except Exception as e:
                print(f"Ошибка обработки письма: {str(e)}")
                continue

        mail.close()
        mail.logout()
    except Exception as e:
        print(f"Ошибка IMAP: {str(e)}")


def update_status(task, status):
    try:
        df = pd.read_excel(EXCEL_FILE)
        # Ищем точное совпадение задачи
        mask = df["Задача"].str.lower() == task
        if not any(mask):
            print(f" Задача {task} не найдена в tasks.xlsx")
            return
            
        df.loc[mask, "Статус"] = status
        df.to_excel(EXCEL_FILE, index=False)
        print(f" Обновлен статус для {task}: {status}")
    except Exception as e:
        print(f" Ошибка при обновлении статуса: {e}")
