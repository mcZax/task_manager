import pandas as pd
from config import EXCEL_FILE, DATE_LOG


def init_log():
    log_columns = ["Задача", "Дата отправки", "Дата обновления статуса"]
    log_df = pd.DataFrame(columns=log_columns)
    log_df.to_excel(DATE_LOG, index=False)

 
def load_tasks():
    try:
        return pd.read_excel(EXCEL_FILE)
    except FileNotFoundError:
        print("Файл tasks.xlsx не найден, создаю новый...")
        df = pd.DataFrame(columns=["Задача", "Исполнитель", "Email", "Дедлайн", "Статус"])
        df.to_excel(EXCEL_FILE, index=False)
        return df
    

def update_status(email, status):
    try:
        df = pd.read_excel(EXCEL_FILE)
        # Ищем точное совпадение email (без учета регистра)
        mask = df["Email"].str.lower() == email.lower()
        if not any(mask):
            print(f" Email {email} не найден в tasks.xlsx")
            return
            
        df.loc[mask, "Статус"] = status
        df.to_excel(EXCEL_FILE, index=False)
        print(f" Обновлен статус для {email}: {status}")
    except Exception as e:
        print(f" Ошибка при обновлении статуса: {e}")