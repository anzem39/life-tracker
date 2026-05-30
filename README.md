# Life Tracker

Трекер привычек, финансов и целей в Google Sheets. Запускаешь один раз — таблица создаётся сама, Apps Script встаёт сам, меню появляется само.

Устал настраивать руками каждый раз — написал скрипт который делает всё за тебя.

## Запуск

```
pip install gspread google-auth google-api-python-client
python tracker_setup.py
```

Нужен `credentials.json` от Google Service Account. Как его получить — в `tracker_README.txt`.
