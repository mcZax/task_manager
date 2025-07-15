import pandas as pd
import smtplib
import time
from datetime import datetime, timedelta
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from config import SMTP_SERVER, SMTP_PORT, EMAIL, PASSWORD, DATE_LOG
from database.excel_handler import load_tasks, init_log
import re


def log_sent_task(task):
    try:
        log_df = init_log()
        
        # Проверяем, существует ли уже такая задача
        task_exists = log_df["Задача"].eq(task).any()
        
        if task_exists:
            # Обновляем время для существующей задачи
            log_df.loc[log_df["Задача"] == task, 
                      "Дата и время напоминания"] = datetime.now().strftime("%Y-%m-%d %H:%M")
        else:
            # Добавляем новую запись
            new_entry = {
                "Задача": task,
                "Дата и время напоминания": datetime.now().strftime("%Y-%m-%d %H:%M"),
                "Дата получения ответа": None,
            }
            log_df = pd.concat([log_df, pd.DataFrame([new_entry])], ignore_index=True)
        
        # Сохраняем изменения
        log_df.to_excel(DATE_LOG, index=False)
        
    except Exception as e:
        print(f"Ошибка при логировании задачи: {e}")
        raise


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

        log_sent_task(task)
        print('Дата и время отправки записано в лог')

    except Exception as e:
        print(f"Ошибка при отправке письма: {e}")


def send_monthly_report():
    today = datetime.now().date()
    
    # Проверяем, что сегодня последний день месяца
    next_day = today + timedelta(days=1)
    # if next_day.month == today.month:
    if today != today:
        print("Сегодня не последний день месяца. Рассылка не требуется.")
        return
    
    print("Подготовка месячных отчетов...")
    
    # Загружаем данные
    try:
        df = load_tasks()
    except Exception as e:
        print(f"Ошибка загрузки задач: {e}")
        return
    
    # Группируем задачи по исполнителям
    grouped = df.groupby(['Email', 'Исполнитель'])
    
    # Настройки SMTP
    smtp_server = SMTP_SERVER
    smtp_port = SMTP_PORT
    smtp_login = EMAIL
    smtp_password = PASSWORD  
    
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
