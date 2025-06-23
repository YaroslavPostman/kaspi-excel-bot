
import os
import re
import requests
import pandas as pd
from collections import Counter
from datetime import datetime
import pytz

# Настройки Telegram
TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN")
TELEGRAM_CHAT_ID = os.getenv("TELEGRAM_CHAT_ID")

# URL для скачивания Excel-файла заказов (не API, а страница Kaspi с заказами)
KASPI_EXCEL_URL = "https://kaspi.kz/shop/api/v2/orders/export?status=KASPI_DELIVERY_CARGO_ASSEMBLY"

# Заголовки для авторизации
HEADERS = {
    "X-Mc-Api-Session-Id": "Y4-1c639cc5-583e-47f7-98ac-e6f5e0f80aac",
    "Cookie": (
        "kaspi.storefront.cookie.city=750000000; "
        "mc-sid=1d86f21b-fa72-4f74-bf9f-67847b5eccdd; "
        "ssaid=79a7e9b0-998d-11ee-9a95-77ef56f38499; "
        "ks.tg=15"
    )
}

def download_excel(path="orders.xlsx"):
    response = requests.get(KASPI_EXCEL_URL, headers=HEADERS)
    if response.status_code == 200:
        with open(path, "wb") as f:
            f.write(response.content)
        return path
    else:
        raise Exception(f"Ошибка загрузки Excel: {response.status_code}")

def extract_sizes_from_excel(path):
    df = pd.read_excel(path)
    titles = df["Название в системе продавца"].dropna()
    valid_sizes = ["XS", "S", "M", "L", "XL", "XXL", "XXXL"]

    def extract_size(title):
        match = re.search(r"черный\s+(XS|S|M|L|XL|XXL|XXXL)", title, re.IGNORECASE)
        return match.group(1).upper() if match else None

    sizes = [extract_size(t) for t in titles]
    sizes = [s for s in sizes if s in valid_sizes]
    return Counter(sizes)

def format_message(counter):
    if not counter:
        return "❌ Нет заказов на сборку."

    lines = ["📦 Заказы на сборку:"]
    for size in ["XS", "S", "M", "L", "XL", "XXL", "XXXL"]:
        if size in counter:
            lines.append(f"{size} – {counter[size]} шт.")
    return "\n".join(lines)

def send_telegram(text):
    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
    payload = {
        "chat_id": TELEGRAM_CHAT_ID,
        "text": text
    }
    requests.post(url, data=payload)

def main():
    print("🚀 KASPI BOT ЗАПУЩЕН")
    try:
        path = download_excel()
        counter = extract_sizes_from_excel(path)
        message = format_message(counter)
    except Exception as e:
        message = f"❌ Ошибка: {e}"

    tz = pytz.timezone('Asia/Almaty')
    time_stamp = datetime.now(tz).strftime("%Y-%m-%d %H:%M:%S")
    message += f"\n🕒 {time_stamp}"
    send_telegram(message)

if __name__ == "__main__":
    main()
