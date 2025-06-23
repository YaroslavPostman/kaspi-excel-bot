import os
import datetime
import pytz
import requests
import pandas as pd

# Телеграм параметры
BOT_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN")
CHAT_ID = os.environ.get("TELEGRAM_CHAT_ID")

def log(msg):
    print(f"📍 {msg}", flush=True)

def send_telegram_message(message):
    log("Отправляем сообщение в Telegram...")
    url = f"https://api.telegram.org/bot{BOT_TOKEN}/sendMessage"
    payload = {
        "chat_id": CHAT_ID,
        "text": message
    }
    response = requests.post(url, data=payload)
    log(f"📬 Telegram response: {response.status_code} {response.text}")

def extract_sizes_from_excel(file_path):
    log("📄 Читаем Excel-файл...")
    df = pd.read_excel(file_path, engine='openpyxl')
    column_c = df.iloc[:, 2].dropna().astype(str)

    sizes = ["XS", "S", "M", "L", "XL", "XXL", "XXXL"]
    count = {size: 0 for size in sizes}

    for name in column_c:
        for size in sizes:
            if name.strip().upper().endswith(f"ЧЕРНЫЙ {size}"):
                count[size] += 1

    return {size: qty for size, qty in count.items() if qty > 0}

def main():
    log("🚀 KASPI Excel BOT запущен")
    file_path = "ActiveOrders.xlsx"

    # ↓↓↓ Здесь вставь ссылку на Excel, если хочешь скачать с Kaspi
    # response = requests.get("https://...", headers=...)
    # with open(file_path, "wb") as f:
    #     f.write(response.content)

    if not os.path.exists(file_path):
        send_telegram_message("❌ Файл ActiveOrders.xlsx не найден.")
        return

    sizes_count = extract_sizes_from_excel(file_path)

    if sizes_count:
        msg = "📦 Заказы на сборку:\n" + "\n".join(f"{k} – {v} шт." for k, v in sizes_count.items())
    else:
        tz = pytz.timezone("Asia/Almaty")
        now = datetime.datetime.now(tz).strftime("%Y-%m-%d %H:%M:%S")
        msg = f"❌ Нет заказов на сборку. Время: {now}"

    send_telegram_message(msg)

if __name__ == "__main__":
    main()
