from mail.sender import send_monthly_report
from mail.parser import check_deadlines, check_responses

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