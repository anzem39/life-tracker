============================================================
  ТРЕКЕР ЖИЗНИ — ИНСТРУКЦИЯ ПО ЗАПУСКУ
============================================================

ШАГ 1: Получить credentials.json
----------------------------------
1. Открой https://console.cloud.google.com/
2. Создай проект (или выбери существующий)
3. Включи Google Sheets API и Google Drive API:
   APIs & Services → Enable APIs → найди и включи оба
4. Создай Service Account:
   APIs & Services → Credentials → Create Credentials → Service Account
   Дай любое имя → Create and Continue → Done
5. Зайди в созданный аккаунт → Keys → Add Key → JSON
   Скачается файл — переименуй в credentials.json
6. Положи credentials.json рядом с tracker_setup.py

ШАГ 2: Установить зависимости
-------------------------------
Открой терминал в папке с файлом и выполни:

  pip install gspread google-auth google-api-python-client

ШАГ 3: Запустить скрипт
-------------------------
  python tracker_setup.py

Скрипт создаст таблицу и выведет ссылку.
Также создастся файл apps_script.gs

ШАГ 4: Установить Apps Script
-------------------------------
1. Открой таблицу по ссылке
2. Расширения → Apps Script
3. Удали весь дефолтный код
4. Вставь содержимое файла apps_script.gs
5. Нажми Сохранить (Ctrl+S)
6. Выбери функцию setupTriggers → нажми Запустить
7. Разреши доступ при запросе (один раз)
8. В таблице появится меню "📊 Трекер"

============================================================
ВОЗМОЖНЫЕ ОШИБКИ
============================================================

"Permission denied" / "Request had insufficient auth scopes"
→ Убедись что в Google Cloud включены оба API (Sheets + Drive)

"File not found: credentials.json"
→ Положи credentials.json в ту же папку что tracker_setup.py

"gspread не найден"
→ Выполни: pip install gspread google-auth google-api-python-client

Таблица не открывается
→ Скрипт даёт доступ "anyone can edit" — просто открой ссылку
============================================================
