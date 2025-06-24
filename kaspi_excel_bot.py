
import os
import datetime
import pytz
import requests
import pandas as pd

# === Настройки Telegram ===
BOT_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN")
CHAT_ID = os.environ.get("TELEGRAM_CHAT_ID")

# === Настройки Kaspi ===
DOWNLOAD_URL = "https://mc.shop.kaspi.kz/order/view/mc/order/export?presetFilter=KASPI_DELIVERY_CARGO_ASSEMBLY&merchantId=30067732&fromDate=1750705200000&toDate=1750791600000&_m=30067732"

COOKIES = {
    "X-Mc-Api-Session-Id": "Y4-1c639cc5-583e-47f7-98ac-e6f5e0f80aac",
    "kaspi.storefront.cookie.city": "750000000",
    "mc-sid": "1d86f21b-fa72-4f74-bf9f-67847b5eccdd",
    "ssaid": "79a7e9b0-998d-11ee-9a95-77ef56f38499",
    "ks.tg": "15"
}

HEADERS = {
    "User-Agent": "Mozilla/5.0",
    "Accept": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
}

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

def download_excel():
    log("⬇️ Скачиваем Excel-файл с Kaspi...")
    response = requests.get(DOWNLOAD_URL, headers=HEADERS, cookies=COOKIES)
    if response.status_code == 200:
        with open("ActiveOrders.xlsx", "wb") as f:
            f.write(response.content)
        log("✅ Файл сохранён как ActiveOrders.xlsx")
        return True
    else:
        log(f"❌ Ошибка загрузки файла: {response.status_code}")
        return False

def extract_sizes_from_excel(file_path):
    log("📄 Читаем Excel-файл...")
    df = pd.read_excel(file_path, engine='openpyxl')

    df = df[df["Статус"] == "Ожидает передачи курьеру"]

    names = df["Название товара в Kaspi Магазине"].dropna().astype(str)

    sizes = ["XS", "S", "M", "L", "XL", "XXL", "XXXL"]
    colors = ["ЧЕРНЫЙ", "БЕЛЫЙ", "СИНИЙ", "КРАСНЫЙ", "БОРДОВЫЙ"]
    result = {}

    for name in names:
        for color in colors:
            if color in name.upper():
                for size in sizes:
                    if name.upper().endswith(f"{color} {size}"):
                        result.setdefault(color.capitalize(), {})
                        result[color.capitalize()].setdefault(size, 0)
                        result[color.capitalize()][size] += 1
    return result

def build_message(stats):
    if not stats:
        tz = pytz.timezone("Asia/Almaty")
        now = datetime.datetime.now(tz).strftime("%Y-%m-%d %H:%M:%S")
        return f"❌ Нет заказов на сборку. Время: {now}"

    lines = ["📦 Заказы на сборку:"]
    for color, sizes in stats.items():
        lines.append(f"
{color}:")
        for size, count in sizes.items():
            lines.append(f"  {size} – {count}")
    return "\n".join(lines)

def main():
    log("🚀 KASPI BOT с автоскачиванием запущен")

    if not download_excel():
        send_telegram_message("❌ Не удалось скачать Excel с Kaspi")
        return

    stats = extract_sizes_from_excel("ActiveOrders.xlsx")
    message = build_message(stats)
    send_telegram_message(message)

if __name__ == "__main__":
    main()
