import sys
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
sys.stderr.reconfigure(encoding='utf-8', errors='replace')

import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import datetime

SERVICE_ACCOUNT_FILE = 'credentials.json'
SPREADSHEET_ID = '1ZQWPZFx2kuARTfady9TvZvsuXGcEuseoRV2x9tV3TSY'

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive',
]

# Цвета тёмной темы (0.0–1.0 для Google API)
BG_DARK       = {'red': 0.102, 'green': 0.102, 'blue': 0.180}  # #1a1a2e
BG_CARD       = {'red': 0.149, 'green': 0.149, 'blue': 0.247}  # #262640
ACCENT_GREEN  = {'red': 0.000, 'green': 0.831, 'blue': 0.667}  # #00d4aa
ACCENT_RED    = {'red': 0.914, 'green': 0.271, 'blue': 0.376}  # #e94560
ACCENT_YELLOW = {'red': 1.000, 'green': 0.843, 'blue': 0.000}  # #FFD700
TEXT_WHITE    = {'red': 1.000, 'green': 1.000, 'blue': 1.000}
TEXT_GREY     = {'red': 0.700, 'green': 0.700, 'blue': 0.750}
WEEKLY_BG     = {'red': 0.149, 'green': 0.118, 'blue': 0.247}

# УТИЛИТЫ
def auth():
    creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    gc = gspread.authorize(creds)
    sheets_svc = build('sheets', 'v4', credentials=creds)
    return gc, sheets_svc


def fmt(bg=None, fg=None, bold=False, size=10, ha='LEFT', va='MIDDLE', wrap='OVERFLOW_CELL'):
    f = {
        'textFormat': {'bold': bold, 'fontSize': size, 'foregroundColor': fg or TEXT_WHITE},
        'horizontalAlignment': ha,
        'verticalAlignment': va,
        'wrapStrategy': wrap,
        'padding': {'top': 3, 'bottom': 3, 'left': 5, 'right': 5},
    }
    if bg:
        f['backgroundColor'] = bg
    return f


def rng(shid, r1, c1, r2, c2):
    return {'sheetId': shid, 'startRowIndex': r1, 'endRowIndex': r2,
            'startColumnIndex': c1, 'endColumnIndex': c2}


def rc(shid, r1, c1, r2, c2, cell_fmt):
    """repeatCell"""
    return {'repeatCell': {'range': rng(shid, r1, c1, r2, c2),
                           'cell': {'userEnteredFormat': cell_fmt},
                           'fields': 'userEnteredFormat'}}


def mg(shid, r1, c1, r2, c2):
    """mergeCells"""
    return {'mergeCells': {'range': rng(shid, r1, c1, r2, c2), 'mergeType': 'MERGE_ALL'}}


def cw(shid, ci, ce, px):
    """column width"""
    return {'updateDimensionProperties': {
        'range': {'sheetId': shid, 'dimension': 'COLUMNS', 'startIndex': ci, 'endIndex': ce},
        'properties': {'pixelSize': px}, 'fields': 'pixelSize'}}


def rh(shid, ri, re, px):
    """row height"""
    return {'updateDimensionProperties': {
        'range': {'sheetId': shid, 'dimension': 'ROWS', 'startIndex': ri, 'endIndex': re},
        'properties': {'pixelSize': px}, 'fields': 'pixelSize'}}


def freeze(shid, rows=1, cols=0, hide_grid=True):
    return {'updateSheetProperties': {
        'properties': {'sheetId': shid, 'gridProperties': {
            'hideGridlines': hide_grid,
            'frozenRowCount': rows,
            'frozenColumnCount': cols,
        }},
        'fields': 'gridProperties.hideGridlines,gridProperties.frozenRowCount,gridProperties.frozenColumnCount'}}


def checkbox(shid, r1, c1, r2, c2):
    return {'setDataValidation': {'range': rng(shid, r1, c1, r2, c2),
            'rule': {'condition': {'type': 'BOOLEAN'}, 'showCustomUi': True}}}


def dropdown(shid, r1, c1, r2, c2, options):
    return {'setDataValidation': {'range': rng(shid, r1, c1, r2, c2),
            'rule': {'condition': {'type': 'ONE_OF_LIST',
                     'values': [{'userEnteredValue': o} for o in options]},
                     'showCustomUi': True}}}


def cond(shid, r1, c1, r2, c2, formula, bg_color):
    return {'addConditionalFormatRule': {'rule': {
        'ranges': [rng(shid, r1, c1, r2, c2)],
        'booleanRule': {
            'condition': {'type': 'CUSTOM_FORMULA', 'values': [{'userEnteredValue': formula}]},
            'format': {'backgroundColor': bg_color}
        }
    }, 'index': 0}}


def expand_cols(svc, sid, shid, count=42):
    """Расширяем лист до нужного количества столбцов"""
    svc.spreadsheets().batchUpdate(spreadsheetId=sid, body={'requests': [{
        'updateSheetProperties': {
            'properties': {'sheetId': shid, 'gridProperties': {'columnCount': count}},
            'fields': 'gridProperties.columnCount'
        }
    }]}).execute()


def clear_cond_rules(svc, sid, shid):
    """Удаляем существующие правила условного форматирования (чтобы не дублировались при повторном запуске)"""
    resp = svc.spreadsheets().get(
        spreadsheetId=sid,
        fields='sheets(properties/sheetId,conditionalFormats)'
    ).execute()
    for sheet in resp.get('sheets', []):
        if sheet['properties']['sheetId'] == shid:
            rules = sheet.get('conditionalFormats', [])
            if rules:
                reqs = [{'deleteConditionalFormatRule': {'sheetId': shid, 'index': i}}
                        for i in range(len(rules) - 1, -1, -1)]
                svc.spreadsheets().batchUpdate(spreadsheetId=sid, body={'requests': reqs}).execute()
            break


def batch(svc, sid, reqs):
    if reqs:
        svc.spreadsheets().batchUpdate(spreadsheetId=sid, body={'requests': reqs}).execute()


# ПОДКЛЮЧЕНИЕ К ТАБЛИЦЕ
def open_and_init_sheets(gc, svc, sid):
    ss = gc.open_by_key(sid)
    existing = {s.title for s in ss.worksheets()}
    needed = ['🏠 Главная', '✅ Привычки', '💰 Финансы', '🎯 Цели',
              '📅 Планер', '📜 История', '📖 Инструкция']

    reqs = []
    if ss.sheet1.title not in needed:
        reqs.append({'updateSheetProperties': {
            'properties': {'sheetId': ss.sheet1.id, 'title': '🏠 Главная', 'index': 0},
            'fields': 'title,index'}})
        existing.discard(ss.sheet1.title)
        existing.add('🏠 Главная')

    for i, title in enumerate(needed[1:], 1):
        if title not in existing:
            reqs.append({'addSheet': {'properties': {'title': title, 'index': i}}})

    if reqs:
        svc.spreadsheets().batchUpdate(spreadsheetId=sid, body={'requests': reqs}).execute()

    ss = gc.open_by_key(sid)  # обновляем объект после изменений
    shids = {s['properties']['title']: s['properties']['sheetId']
             for s in svc.spreadsheets().get(spreadsheetId=sid).execute()['sheets']}
    ws = {name: ss.worksheet(name) for name in shids}
    return ws, shids


# ЛИСТ: ГЛАВНАЯ
def setup_glavnaya(ws, shid, svc, sid):
    now = datetime.datetime.now()
    month_ru = ['Январь','Февраль','Март','Апрель','Май','Июнь',
                'Июль','Август','Сентябрь','Октябрь','Ноябрь','Декабрь'][now.month - 1]
    quotes = [
        'Дисциплина — это мост между целями и достижениями.',
        'Каждый день — новый шанс стать лучше.',
        'Маленькие шаги ведут к большим результатам.',
        'Твои привычки формируют твою судьбу.',
        'Сегодняшние усилия — завтрашние результаты.',
    ]
    q = '=CHOOSE(MOD(DAY(TODAY()),5)+1,"{}","{}","{}","{}","{}") '.format(*quotes)

    # % за сегодня — берём из строки 15 (% за день) нужный столбец
    today_pct = "=IFERROR(ROUND(INDEX('✅ Привычки'!B15:AF15,1,DAY(TODAY())),0)&\"%\",\"—\")"

    data = [
        [''] * 8,
        ['', f'📊 ТРЕКЕР ЖИЗНИ — {month_ru.upper()} {now.year}', '', '', '', '', '', ''],
        [''] * 8,
        ['', '📈 ВЫПОЛНЕНИЕ СЕГОДНЯ', '', '', '📅 СТАТИСТИКА МЕСЯЦА', '', '', ''],
        [''] * 8,
        ['', today_pct, '', '', 'Дней прошло:', '', '=DAY(TODAY())', ''],
        ['', '% привычек за сегодня', '', '', 'Средняя оценка:', '',
         "=IFERROR(TEXT(AVERAGE('✅ Привычки'!B12:AF12),\"0.0\"),\"—\")", ''],
        ['', '', '', '', 'Лучший streak:', '',
         "=IFERROR(MAX('✅ Привычки'!AJ2:AJ14),\"—\")", ''],
        [''] * 8,
        ['', '💡 МОТИВАЦИЯ ДНЯ', '', '', '', '', '', ''],
        ['', q, '', '', '', '', '', ''],
        [''] * 8,
        ['', '🔥 СТРИКИ ПРИВЫЧЕК', '', '', '', '', '', ''],
        ['', '🚿 Холодный душ',         '', "=IFERROR('✅ Привычки'!AJ2,0)", '', '', '', ''],
        ['', '💧 3л воды',               '', "=IFERROR('✅ Привычки'!AJ3,0)", '', '', '', ''],
        ['', '🚫 Без алкоголя',          '', "=IFERROR('✅ Привычки'!AJ4,0)", '', '', '', ''],
        ['', '📵 Без порнохаба',         '', "=IFERROR('✅ Привычки'!AJ5,0)", '', '', '', ''],
        ['', '📚 Чтение 30 мин',         '', "=IFERROR('✅ Привычки'!AJ6,0)", '', '', '', ''],
        ['', '💊 Антидепрессанты',       '', "=IFERROR('✅ Привычки'!AJ7,0)", '', '', '', ''],
        ['', '🔇 Без соцсетей до 12:00', '', "=IFERROR('✅ Привычки'!AJ8,0)", '', '', '', ''],
        ['', '📝 Планы на завтра',       '', "=IFERROR('✅ Привычки'!AJ9,0)", '', '', '', ''],
        ['', '🧹 Уборка комнаты',        '', "=IFERROR('✅ Привычки'!AJ14,0)", '', '', '', ''],
        [''] * 8,
        ['', '⭐ ТОП ПРИВЫЧКА НЕДЕЛИ', '', '', '', '', '', ''],
        ['', "=IFERROR(INDEX('✅ Привычки'!A2:A14,MATCH(MAX('✅ Привычки'!AK2:AK14),'✅ Привычки'!AK2:AK14,0)),\"—\")",
         '', '', '', '', '', ''],
        [''] * 8,
    ]
    ws.update(data, 'A1', value_input_option='USER_ENTERED')

    reqs = [
        rc(shid, 0, 0, 30, 10, fmt(bg=BG_DARK)),
        # Заголовок
        mg(shid, 1, 1, 2, 7),
        rc(shid, 1, 1, 2, 7, fmt(bg=BG_DARK, fg=ACCENT_GREEN, bold=True, size=20, ha='CENTER')),
        # Секция "Сегодня"
        mg(shid, 3, 1, 4, 4),
        rc(shid, 3, 1, 4, 4, fmt(bg=BG_DARK, fg=ACCENT_GREEN, bold=True, size=12)),
        mg(shid, 4, 1, 8, 4),
        rc(shid, 4, 1, 8, 4, fmt(bg=BG_CARD, fg=ACCENT_GREEN, bold=True, size=40, ha='CENTER', va='MIDDLE')),
        # Секция статистики
        mg(shid, 3, 4, 4, 7),
        rc(shid, 3, 4, 4, 7, fmt(bg=BG_DARK, fg=ACCENT_GREEN, bold=True, size=12)),
        rc(shid, 4, 4, 8, 5, fmt(bg=BG_CARD, fg=TEXT_GREY, size=11)),
        rc(shid, 4, 6, 8, 7, fmt(bg=BG_CARD, fg=ACCENT_GREEN, bold=True, size=13, ha='CENTER')),
        # Мотивация
        mg(shid, 9, 1, 10, 7),
        rc(shid, 9, 1, 10, 7, fmt(bg=BG_DARK, fg=ACCENT_GREEN, bold=True, size=12)),
        mg(shid, 10, 1, 11, 7),
        rc(shid, 10, 1, 11, 7, fmt(bg=BG_CARD, fg=TEXT_WHITE, size=12, ha='CENTER', wrap='WRAP')),
        # Стрики заголовок
        mg(shid, 12, 1, 13, 7),
        rc(shid, 12, 1, 13, 7, fmt(bg=BG_DARK, fg=ACCENT_GREEN, bold=True, size=12)),
        # Топ привычка
        mg(shid, 23, 1, 24, 7),
        rc(shid, 23, 1, 24, 7, fmt(bg=BG_DARK, fg=ACCENT_GREEN, bold=True, size=12)),
        mg(shid, 24, 1, 25, 7),
        rc(shid, 24, 1, 25, 7, fmt(bg=BG_CARD, fg=ACCENT_YELLOW, bold=True, size=14, ha='CENTER')),
        # Ширины
        cw(shid, 0, 1, 20), cw(shid, 1, 2, 240), cw(shid, 2, 3, 10),
        cw(shid, 3, 4, 70), cw(shid, 4, 5, 180), cw(shid, 5, 6, 10), cw(shid, 6, 7, 80),
        rh(shid, 4, 8, 55), rh(shid, 10, 11, 50),
        freeze(shid, rows=1, hide_grid=True),
    ]
    # Строки стриков (индексы 13-21)
    for i in range(13, 22):
        reqs += [
            rc(shid, i, 1, i+1, 2, fmt(bg=BG_CARD, fg=TEXT_GREY, size=11)),
            rc(shid, i, 3, i+1, 4, fmt(bg=BG_CARD, fg=ACCENT_GREEN, bold=True, size=16, ha='CENTER')),
        ]
    batch(svc, sid, reqs)


# ЛИСТ: ПРИВЫЧКИ
# Структура строк в листе (1-based / 0-based index):
#  1/0  : заголовок
#  2-9 / 1-8  : ежедневные чекбоксы (HABITS 0-7)
#  10/9        : Часы работы (number)
#  11/10       : Задачи выполнены (dropdown)
#  12/11       : Оценка дня (number)
#  13/12       : РАЗДЕЛИТЕЛЬ (пустая строка)
#  14/13       : Уборка комнаты (еженедельный чекбокс)
#  15/14       : % за день (формула)
#  16/15       : Заметка дня (текст)

HABITS = [
    ('🚿 Холодный душ',           'checkbox'),
    ('💧 3л воды',                 'checkbox'),
    ('🚫 Без алкоголя',            'checkbox'),
    ('📵 Без порнохаба',           'checkbox'),
    ('📚 Чтение 30 мин',           'checkbox'),
    ('💊 Антидепрессанты 09-15',   'checkbox'),
    ('🔇 Без соцсетей до 12:00',   'checkbox'),
    ('📝 Планы на завтра',         'checkbox'),
    ('💼 Часы работы',             'number'),
    ('✅ Задачи выполнены',         'dropdown'),
    ('⭐ Оценка дня (1-10)',        'number'),
    ('', None),                              # разделитель → строка 13 (index 12)
    ('🧹 Уборка комнаты (еженед.)', 'checkbox'),
]


def setup_privychki(ws, shid, svc, sid):
    expand_cols(svc, sid, shid, 42)

    header = ['Привычка'] + [str(d) for d in range(1, 32)] + ['🔥 Streak', '📅 % мес.', '📊']
    rows = [header]
    for name, _ in HABITS:
        rows.append([name] + [''] * 31 + ['', '', ''] if name else [''] * len(header))

    # Строка % за день: считаем чекбоксы строк 2-9 + строка 14 (уборка)
    pct_row = ['% за день']
    for d in range(1, 32):
        pct_row.append(
            f'=IFERROR(('
            f'COUNTIF(INDIRECT(ADDRESS(2,{d+1})&":"&ADDRESS(9,{d+1})),TRUE)+'
            f'IF(INDIRECT(ADDRESS(14,{d+1}))=TRUE,1,0)'
            f')/9*100,0)'
        )
    pct_row += ['', '', '']
    rows.append(pct_row)
    rows.append(['✏️ Заметка дня'] + [''] * 31 + ['', '', ''])

    ws.update(rows, 'A1', value_input_option='USER_ENTERED')

    # Streak (AJ) и % за месяц (AK) для чекбоксов
    for i, (name, kind) in enumerate(HABITS):
        if not name or kind != 'checkbox':
            continue
        row = i + 2  # 1-based
        ws.update(
            [[f'=IFERROR(COUNTIF(INDIRECT("B{row}:"&ADDRESS({row},MIN(DAY(TODAY())+1,32))),TRUE),0)']],
            f'AJ{row}', value_input_option='USER_ENTERED')
        ws.update(
            [[f'=IFERROR(COUNTIF(B{row}:AF{row},TRUE)/MAX(DAY(TODAY()),1)*100,0)']],
            f'AK{row}', value_input_option='USER_ENTERED')

    ws.update([['=IFERROR(AVERAGE(B10:AF10),0)']], 'AK10', value_input_option='USER_ENTERED')
    ws.update([['=IFERROR(AVERAGE(B12:AF12),0)']], 'AK12', value_input_option='USER_ENTERED')

    clear_cond_rules(svc, sid, shid)

    reqs = [
        rc(shid, 0, 0, 20, 42, fmt(bg=BG_DARK, fg=TEXT_WHITE, size=10)),
        # Заголовок
        rc(shid, 0, 0, 1, 1,  fmt(bg=BG_CARD, fg=TEXT_WHITE, bold=True, size=11)),
        rc(shid, 0, 1, 1, 32, fmt(bg=BG_CARD, fg=TEXT_WHITE, bold=True, size=11, ha='CENTER')),
        rc(shid, 0, 35, 1, 38, fmt(bg=BG_CARD, fg=ACCENT_GREEN, bold=True, size=11, ha='CENTER')),
        # Ежедневные чекбоксы строки 2-9 (индексы 1-8)
        rc(shid, 1, 0, 9, 1,  fmt(bg=BG_CARD, fg=TEXT_WHITE, size=11)),
        rc(shid, 1, 1, 9, 32, fmt(bg=BG_DARK,  fg=TEXT_WHITE, size=10, ha='CENTER')),
        rc(shid, 1, 35, 9, 38, fmt(bg=BG_CARD, fg=ACCENT_GREEN, bold=True, size=11, ha='CENTER')),
        # Числовые строки 10-12 (индексы 9-11)
        rc(shid, 9, 0, 12, 1,  fmt(bg=BG_CARD, fg=TEXT_WHITE, size=11)),
        rc(shid, 9, 1, 12, 32, fmt(bg=BG_DARK,  fg=TEXT_WHITE, size=10, ha='CENTER')),
        rc(shid, 9, 35, 12, 38, fmt(bg=BG_CARD, fg=ACCENT_GREEN, size=11, ha='CENTER')),
        # Разделитель строка 13 (индекс 12) — ПРАВИЛЬНЫЙ индекс
        rc(shid, 12, 0, 13, 38, fmt(bg=BG_DARK)),
        # Еженедельная привычка строка 14 (индекс 13)
        rc(shid, 13, 0, 14, 1,  fmt(bg=WEEKLY_BG, fg=ACCENT_GREEN, bold=True, size=11)),
        rc(shid, 13, 1, 14, 32, fmt(bg=WEEKLY_BG, fg=TEXT_WHITE, size=10, ha='CENTER')),
        rc(shid, 13, 35, 14, 38, fmt(bg=WEEKLY_BG, fg=ACCENT_GREEN, bold=True, size=11, ha='CENTER')),
        # % за день строка 15 (индекс 14)
        rc(shid, 14, 0, 15, 1,  fmt(bg=BG_CARD, fg=ACCENT_GREEN, bold=True, size=11)),
        rc(shid, 14, 1, 15, 32, fmt(bg=BG_CARD, fg=ACCENT_GREEN, bold=True, size=11, ha='CENTER')),
        # Заметка строка 16 (индекс 15)
        rc(shid, 15, 0, 16, 1,  fmt(bg=BG_CARD, fg=TEXT_GREY, bold=True, size=11)),
        rc(shid, 15, 1, 16, 32, fmt(bg=BG_DARK,  fg=TEXT_WHITE, size=10, wrap='WRAP')),
        # Валидации
        checkbox(shid, 1, 1, 9, 32),
        checkbox(shid, 13, 1, 14, 32),
        dropdown(shid, 9, 1, 10, 32, ['Да', 'Нет', '—']),
        # Условное форматирование (формулы с явной ссылкой на верхний левый угол)
        cond(shid, 1, 1, 9, 32, '=B2=TRUE',
             {'red': 0.0, 'green': 0.45, 'blue': 0.27}),
        cond(shid, 13, 1, 14, 32, '=B14=TRUE',
             {'red': 0.0, 'green': 0.35, 'blue': 0.20}),
        # Размеры
        cw(shid, 0, 1, 230), cw(shid, 1, 32, 36),
        cw(shid, 35, 36, 80), cw(shid, 36, 37, 90), cw(shid, 37, 38, 70),
        rh(shid, 0, 1, 36), rh(shid, 1, 15, 32), rh(shid, 15, 16, 48),
        freeze(shid, rows=1, cols=1, hide_grid=True),
    ]
    batch(svc, sid, reqs)


# ЛИСТ: ФИНАНСЫ
def setup_finansy(ws, shid, svc, sid):
    # Строки данных:
    # 1  : заголовок
    # 2  : пусто
    # 3  : бюджет (C3)
    # 4  : итого (SUM формулы)
    # 5  : остаток
    # 6  : пусто
    # 7  : ДОХОДЫ header
    # 8  : колонки
    # 9-38 : 30 строк доходов (H9:H38)
    # 39-40: пусто
    # 41  : РАСХОДЫ header
    # 42  : колонки
    # 43-102: 60 строк расходов (H43:H102)

    data = [
        ['💰 ФИНАНСЫ'] + [''] * 7,
        [''] * 8,
        ['📊 Бюджет на месяц (₸):', '', 0, '', '', '', '', ''],
        ['Итого доходов:', '', '=SUM(H9:H38)', '', 'Итого расходов:', '', '=SUM(H43:H102)', ''],
        ['Остаток:', '', '=C4-G4', '', '% бюджета:', '', '=IFERROR(TEXT(G4/C3*100,"0.0")&"%","—")', ''],
        [''] * 8,
        ['📈 ДОХОДЫ'] + [''] * 7,
        ['Дата', 'Источник дохода', '', '', '', '', '', 'Сумма ₸'],
    ]
    for _ in range(30):
        data.append([''] * 8)
    data += [[''] * 8, [''] * 8,
             ['📉 РАСХОДЫ'] + [''] * 7,
             ['Дата', 'Категория', '', 'Описание', '', '', '', 'Сумма ₸']]
    for _ in range(60):
        data.append([''] * 8)

    ws.update(data, 'A1', value_input_option='USER_ENTERED')

    reqs = [
        rc(shid, 0, 0, 130, 10, fmt(bg=BG_DARK, fg=TEXT_WHITE, size=11)),
        mg(shid, 0, 0, 1, 7),
        rc(shid, 0, 0, 1, 7, fmt(bg=BG_DARK, fg=ACCENT_GREEN, bold=True, size=20, ha='CENTER')),
        # Бюджет
        rc(shid, 2, 0, 3, 2, fmt(bg=BG_CARD, fg=TEXT_WHITE, bold=True, size=12)),
        rc(shid, 2, 2, 3, 3, fmt(bg=BG_CARD, fg=ACCENT_GREEN, bold=True, size=14, ha='CENTER')),
        # Итоги
        rc(shid, 3, 0, 5, 2, fmt(bg=BG_CARD, fg=TEXT_GREY, bold=True, size=11)),
        rc(shid, 3, 2, 5, 3, fmt(bg=BG_CARD, fg=ACCENT_GREEN, bold=True, size=12, ha='RIGHT')),
        rc(shid, 3, 4, 5, 6, fmt(bg=BG_CARD, fg=TEXT_GREY, bold=True, size=11)),
        rc(shid, 3, 6, 5, 7, fmt(bg=BG_CARD, fg=ACCENT_RED, bold=True, size=12, ha='RIGHT')),
        # Заголовки разделов
        mg(shid, 6, 0, 7, 7),
        rc(shid, 6, 0, 7, 7, fmt(bg=BG_CARD, fg=ACCENT_GREEN, bold=True, size=14)),
        rc(shid, 7, 0, 8, 7, fmt(bg=BG_CARD, fg=TEXT_WHITE, bold=True, size=11, ha='CENTER')),
        mg(shid, 40, 0, 41, 7),
        rc(shid, 40, 0, 41, 7, fmt(bg=BG_CARD, fg=ACCENT_GREEN, bold=True, size=14)),
        rc(shid, 41, 0, 42, 7, fmt(bg=BG_CARD, fg=TEXT_WHITE, bold=True, size=11, ha='CENTER')),
        # Категории расходов (строки 43-102, индексы 42-101)
        dropdown(shid, 42, 1, 102, 2, ['Еда', 'Транспорт', 'Развлечения', 'Подписки', 'Одежда', 'Другое']),
        # Размеры
        cw(shid, 0, 1, 110), cw(shid, 1, 2, 160), cw(shid, 2, 3, 20),
        cw(shid, 3, 4, 220), cw(shid, 4, 7, 20), cw(shid, 7, 8, 140),
        rh(shid, 0, 1, 50), rh(shid, 2, 6, 36),
        freeze(shid, rows=1, hide_grid=True),
    ]
    batch(svc, sid, reqs)


# ЛИСТ: ЦЕЛИ
def setup_tseli(ws, shid, svc, sid):
    header = ['Название цели', 'Тип', 'Дедлайн', 'Прогресс %', 'Прогресс-бар', 'Статус', 'Заметки']
    data = [
        ['🎯 ЦЕЛИ'] + [''] * 6,
        [''] * 7,
        header,
    ]
    for i in range(10):
        r = i + 4
        bar = f'=REPT("█",ROUND(D{r}/10,0))&REPT("░",10-ROUND(D{r}/10,0))'
        data.append(['', 'Месячная', '', 0, bar, 'В процессе', ''])

    ws.update(data, 'A1', value_input_option='USER_ENTERED')

    clear_cond_rules(svc, sid, shid)
    reqs = [
        rc(shid, 0, 0, 15, 8, fmt(bg=BG_DARK, fg=TEXT_WHITE, size=11)),
        mg(shid, 0, 0, 1, 6),
        rc(shid, 0, 0, 1, 6, fmt(bg=BG_DARK, fg=ACCENT_GREEN, bold=True, size=20, ha='CENTER')),
        rc(shid, 2, 0, 3, 7, fmt(bg=BG_CARD, fg=ACCENT_GREEN, bold=True, size=11, ha='CENTER')),
        rc(shid, 3, 4, 13, 5, fmt(bg=BG_CARD, fg=ACCENT_GREEN, size=10)),
        dropdown(shid, 3, 1, 13, 2, ['Недельная', 'Месячная', 'Годовая']),
        dropdown(shid, 3, 5, 13, 6, ['В процессе', 'Выполнено', 'Провалено']),
        *[rc(shid, r, 0, r+1, 7, fmt(bg=BG_CARD if r % 2 == 1 else BG_DARK, fg=TEXT_WHITE, size=11))
          for r in range(3, 13)],
        cond(shid, 3, 0, 13, 7, '=$F4="Выполнено"', {'red': 0.0, 'green': 0.3, 'blue': 0.18}),
        cond(shid, 3, 0, 13, 7, '=$F4="Провалено"', {'red': 0.4, 'green': 0.08, 'blue': 0.08}),
        cw(shid, 0, 1, 260), cw(shid, 1, 2, 110), cw(shid, 2, 3, 100),
        cw(shid, 3, 4, 90),  cw(shid, 4, 5, 140), cw(shid, 5, 6, 110), cw(shid, 6, 7, 220),
        rh(shid, 0, 1, 50), rh(shid, 3, 13, 36),
        freeze(shid, rows=3, hide_grid=True),
    ]
    batch(svc, sid, reqs)


# ЛИСТ: ПЛАНЕР
def setup_planer(ws, shid, svc, sid):
    days = ['Пн', 'Вт', 'Ср', 'Чт', 'Пт', 'Сб', 'Вс']
    hours = [f'{h}:00' for h in range(8, 24)]

    data = [
        ['📅 НЕДЕЛЬНЫЙ ПЛАНЕР'] + [''] * 7,
        [''] * 8,
        ['⭐ ГЛАВНЫЕ ЗАДАЧИ НЕДЕЛИ'] + [''] * 7,
        ['1.'] + [''] * 7,
        ['2.'] + [''] * 7,
        ['3.'] + [''] * 7,
        [''] * 8,
        ['Время'] + days,
    ]
    for h in hours:
        data.append([h] + [''] * 7)

    итог = len(data)  # индекс строки начала блока "Итог"
    data += [
        [''] * 8,
        ['📋 ИТОГ НЕДЕЛИ'] + [''] * 7,
        ['Оценка недели (1-10):'] + [''] * 7,
        ['Что получилось:'] + [''] * 7,
        ['Что не вышло:'] + [''] * 7,
    ]

    ws.update(data, 'A1', value_input_option='USER_ENTERED')

    reqs = [
        rc(shid, 0, 0, len(data) + 2, 9, fmt(bg=BG_DARK, fg=TEXT_WHITE, size=11)),
        mg(shid, 0, 0, 1, 7),
        rc(shid, 0, 0, 1, 7, fmt(bg=BG_DARK, fg=ACCENT_GREEN, bold=True, size=20, ha='CENTER')),
        mg(shid, 2, 0, 3, 7),
        rc(shid, 2, 0, 3, 7, fmt(bg=BG_CARD, fg=ACCENT_GREEN, bold=True, size=13)),
        *[mg(shid, r, 1, r+1, 7) for r in range(3, 6)],
        *[rc(shid, r, 0, r+1, 7, fmt(bg=BG_CARD, fg=TEXT_WHITE, bold=True, size=12)) for r in range(3, 6)],
        rc(shid, 7, 0, 8, 8, fmt(bg=BG_CARD, fg=TEXT_WHITE, bold=True, size=12, ha='CENTER')),
        *[rc(shid, 8+i, 0, 9+i, 1, fmt(bg=BG_CARD, fg=TEXT_GREY, bold=True, ha='RIGHT'))
          for i in range(len(hours))],
        *[rc(shid, 8+i, 1, 9+i, 8, fmt(bg=BG_DARK if i % 2 == 0 else BG_CARD, fg=TEXT_WHITE))
          for i in range(len(hours))],
        mg(shid, итог+1, 0, итог+2, 7),
        rc(shid, итог+1, 0, итог+2, 7, fmt(bg=BG_CARD, fg=ACCENT_GREEN, bold=True, size=13)),
        *[mg(shid, итог+2+i, 1, итог+3+i, 7) for i in range(3)],
        *[rc(shid, итог+2+i, 0, итог+3+i, 7, fmt(bg=BG_CARD, fg=TEXT_WHITE, size=11)) for i in range(3)],
        cw(shid, 0, 1, 90),
        *[cw(shid, i, i+1, 145) for i in range(1, 8)],
        rh(shid, 0, 1, 50), rh(shid, 3, 6, 40), rh(shid, 8, 8+len(hours), 28),
        *[rh(shid, итог+2+i, итог+3+i, 50) for i in range(3)],
        freeze(shid, rows=1, hide_grid=True),
    ]
    batch(svc, sid, reqs)


# ЛИСТ: ИСТОРИЯ
def setup_istoriya(ws, shid, svc, sid):
    habits_short = ['🚿 Душ', '💧 Вода', '🚫 Алко', '📵 Порно', '📚 Чтение',
                    '💊 Антидепр.', '🔇 Соцсети', '📝 Планы', '🧹 Уборка']
    header = ['Месяц', 'Дней'] + habits_short + ['⭐ Ср.оценка', '🔥 Макс.streak']
    n = len(header)

    data = [
        ['📜 ИСТОРИЯ ПО МЕСЯЦАМ'] + [''] * (n - 1),
        [''] * n,
        header,
    ]
    ws.update(data, 'A1', value_input_option='USER_ENTERED')

    reqs = [
        rc(shid, 0, 0, 50, n + 1, fmt(bg=BG_DARK, fg=TEXT_WHITE, size=11)),
        mg(shid, 0, 0, 1, n),
        rc(shid, 0, 0, 1, n, fmt(bg=BG_DARK, fg=ACCENT_GREEN, bold=True, size=20, ha='CENTER')),
        rc(shid, 2, 0, 3, n, fmt(bg=BG_CARD, fg=ACCENT_GREEN, bold=True, size=11, ha='CENTER')),
        cw(shid, 0, 1, 160),
        *[cw(shid, i, i+1, 90) for i in range(1, n)],
        rh(shid, 0, 1, 50),
        freeze(shid, rows=3, hide_grid=True),
    ]
    batch(svc, sid, reqs)


# ЛИСТ: ИНСТРУКЦИЯ
def setup_instrukciya(ws, shid, svc, sid):
    lines = [
        ['📖 ИНСТРУКЦИЯ', ''],
        ['', ''],
        ['🚀 НАЧАЛО РАБОТЫ', ''],
        ['1.', 'Откройте лист "✅ Привычки"'],
        ['2.', 'Каждый день отмечайте выполненные привычки чекбоксами'],
        ['3.', 'Вводите часы работы (число), оценку дня (1-10), статус задач (Да/Нет)'],
        ['4.', 'Добавляйте заметку о дне в нижнюю строку каждого дня'],
        ['', ''],
        ['💰 ФИНАНСЫ', ''],
        ['1.', 'Откройте "💰 Финансы" и введите бюджет на месяц в ячейку C3'],
        ['2.', 'Записывайте доходы в раздел ДОХОДЫ, расходы — в РАСХОДЫ'],
        ['3.', 'Для расходов выбирайте категорию из выпадающего списка'],
        ['4.', 'Остаток и % использования считаются автоматически'],
        ['', ''],
        ['🎯 ЦЕЛИ', ''],
        ['1.', 'Введите название цели, тип (недельная/месячная/годовая), дедлайн'],
        ['2.', 'Раз в неделю обновляйте % прогресса вручную'],
        ['3.', 'Меняйте статус: В процессе → Выполнено / Провалено'],
        ['', ''],
        ['📅 ПЛАНЕР', ''],
        ['1.', 'Каждое воскресенье вписывайте 3 главные задачи на неделю'],
        ['2.', 'Заполняйте временные слоты нужными задачами'],
        ['3.', 'В конце недели заполните блок "Итог недели"'],
        ['', ''],
        ['📜 АРХИВАЦИЯ МЕСЯЦА', ''],
        ['1.', 'В конце месяца: меню 📊 Трекер → Новый месяц'],
        ['2.', 'Данные сохранятся в "📜 История", лист Привычки очистится'],
        ['', ''],
        ['⚙️ УСТАНОВКА APPS SCRIPT', ''],
        ['1.', 'Откройте Расширения → Apps Script'],
        ['2.', 'Вставьте содержимое файла apps_script.gs'],
        ['3.', 'Сохраните (Ctrl+S), запустите функцию setupTriggers()'],
        ['4.', 'Разрешите доступ при запросе — один раз'],
        ['5.', 'В таблице появится меню "📊 Трекер"'],
        ['', ''],
        ['🔔 АНТИДЕПРЕССАНТЫ', ''],
        ['—', 'В промежутке 09:00–15:00 если чекбокс не отмечен — ячейка краснеет'],
        ['—', 'После отметки подсветка снимается автоматически'],
        ['', ''],
        ['➕ ДОБАВИТЬ ПРИВЫЧКУ', ''],
        ['1.', 'Вставьте строку в диапазон 2–9 на листе Привычки'],
        ['2.', 'Введите название, добавьте чекбоксы: Данные → Проверка данных → Флажок'],
        ['3.', 'Скопируйте формулы Streak/% из соседней строки'],
    ]

    ws.update(lines, 'A1', value_input_option='USER_ENTERED')

    section_rows = [0, 2, 8, 14, 19, 23, 27, 33, 37]
    reqs = [
        rc(shid, 0, 0, len(lines) + 2, 3, fmt(bg=BG_DARK, fg=TEXT_WHITE, size=11, wrap='WRAP')),
        mg(shid, 0, 0, 1, 1),
        rc(shid, 0, 0, 1, 1, fmt(bg=BG_DARK, fg=ACCENT_GREEN, bold=True, size=20, ha='CENTER')),
        *[mg(shid, r, 0, r+1, 1) for r in section_rows if r != 0],
        *[rc(shid, r, 0, r+1, 1, fmt(bg=BG_CARD, fg=ACCENT_GREEN, bold=True, size=13))
          for r in section_rows if r != 0],
        cw(shid, 0, 1, 40), cw(shid, 1, 2, 620),
        rh(shid, 0, 1, 50),
        freeze(shid, rows=1, hide_grid=True),
    ]
    batch(svc, sid, reqs)


# MAIN
def main():
    print('Запуск...')
    gc, sheets_svc = auth()

    print('Подключаемся к таблице...')
    ws, shids = open_and_init_sheets(gc, sheets_svc, SPREADSHEET_ID)

    steps = [
        ('Главная',    setup_glavnaya,     '🏠 Главная'),
        ('Привычки',   setup_privychki,    '✅ Привычки'),
        ('Финансы',    setup_finansy,      '💰 Финансы'),
        ('Цели',       setup_tseli,        '🎯 Цели'),
        ('Планер',     setup_planer,       '📅 Планер'),
        ('История',    setup_istoriya,     '📜 История'),
        ('Инструкция', setup_instrukciya,  '📖 Инструкция'),
    ]
    for label, fn, sheet_name in steps:
        print(f'{label}...')
        fn(ws[sheet_name], shids[sheet_name], sheets_svc, SPREADSHEET_ID)

    with open('apps_script.gs', 'w', encoding='utf-8') as f:
        f.write(APPS_SCRIPT)

    url = f'https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/edit'
    print()
    print('=' * 60)
    print('ГОТОВО!')
    print(url)
    print()
    print('Следующий шаг: Apps Script')
    print('1. Расширения -> Apps Script')
    print('2. Вставить код из apps_script.gs')
    print('3. Запустить setupTriggers()')
    print('=' * 60)


# APPS SCRIPT
APPS_SCRIPT = r"""
// Трекер жизни — Apps Script
// Вставить в Расширения → Apps Script, затем запустить setupTriggers() один раз

const SH_HABITS  = '✅ Привычки';
const SH_HISTORY = '📜 История';

function setupTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  const ss = SpreadsheetApp.getActive();
  ScriptApp.newTrigger('onOpenHandler').forSpreadsheet(ss).onOpen().create();
  ScriptApp.newTrigger('onEditHandler').forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger('checkMeds').timeBased().everyHours(1).create();
  SpreadsheetApp.getUi().alert('Триггеры установлены!');
}

function onOpenHandler() {
  highlightToday();
  checkMeds();
  SpreadsheetApp.getUi()
    .createMenu('📊 Трекер')
    .addItem('📅 Перейти на сегодня', 'goToToday')
    .addItem('🗓️ Новый месяц',        'newMonth')
    .addItem('⚙️ Триггеры',           'setupTriggers')
    .addToUi();
}

function highlightToday() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SH_HABITS);
  if (!sheet) return;
  const ROWS = 16;
  sheet.getRange(1, 2, ROWS, 31).setBorder(false, false, false, false, false, false);
  const col = new Date().getDate() + 1;
  if (col <= 32) {
    sheet.getRange(1, col, ROWS, 1)
         .setBorder(true, true, true, true, false, false, '#FFD700',
                    SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  }
}

function goToToday() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(SH_HABITS);
  if (!sheet) return;
  ss.setActiveSheet(sheet);
  sheet.setActiveRange(sheet.getRange(2, new Date().getDate() + 1));
}

function onEditHandler(e) {
  const sheet = e.source.getActiveSheet();
  if (sheet.getName() !== SH_HABITS) return;
  const row = e.range.getRow();
  if ((row >= 2 && row <= 9) || row === 14) {
    recalcStreak(sheet, row);
  }
}

function recalcStreak(sheet, row) {
  const today = Math.min(new Date().getDate(), 31);
  let streak = 0;
  for (let d = today; d >= 1; d--) {
    if (sheet.getRange(row, d + 1).getValue() === true) streak++;
    else break;
  }
  sheet.getRange(row, 36).setValue(streak);
}

function checkMeds() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SH_HABITS);
  if (!sheet) return;
  const now = new Date();
  const hour = now.getHours();
  const col = now.getDate() + 1;
  const cell = sheet.getRange(7, col);
  if (hour >= 9 && hour < 15) {
    if (cell.getValue() !== true) {
      cell.setBackground('#e94560');
      cell.setNote('Прими антидепрессанты! (09:00–15:00)');
    } else {
      cell.setBackground(null);
      cell.clearNote();
    }
  } else {
    cell.setBackground(null);
    cell.clearNote();
  }
}

function newMonth() {
  const ui = SpreadsheetApp.getUi();
  const btn = ui.alert('Новый месяц',
    'Архивировать текущий месяц и очистить Привычки?',
    ui.ButtonSet.YES_NO);
  if (btn !== ui.Button.YES) return;

  const ss      = SpreadsheetApp.getActive();
  const habits  = ss.getSheetByName(SH_HABITS);
  const history = ss.getSheetByName(SH_HISTORY);
  if (!habits || !history) { ui.alert('Ошибка: листы не найдены'); return; }

  const now  = new Date();
  const prev = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  const MONTHS = ['Январь','Февраль','Март','Апрель','Май','Июнь',
                  'Июль','Август','Сентябрь','Октябрь','Ноябрь','Декабрь'];
  const label = MONTHS[prev.getMonth()] + ' ' + prev.getFullYear();
  const daysGone = new Date(now.getFullYear(), now.getMonth(), 0).getDate();

  const habitRows = [2, 3, 4, 5, 6, 7, 8, 9, 14];
  const pcts = habitRows.map(r => {
    let cnt = 0;
    for (let d = 1; d <= daysGone; d++) {
      if (habits.getRange(r, d + 1).getValue() === true) cnt++;
    }
    return Math.round(cnt / daysGone * 100) + '%';
  });

  let sum = 0, cnt = 0;
  for (let d = 1; d <= daysGone; d++) {
    const v = habits.getRange(12, d + 1).getValue();
    if (v > 0) { sum += v; cnt++; }
  }
  const avgRating = cnt > 0 ? (sum / cnt).toFixed(1) : '—';

  let maxStreak = 0;
  for (let r = 2; r <= 9; r++) {
    const s = habits.getRange(r, 36).getValue();
    if (s > maxStreak) maxStreak = s;
  }

  const lastRow = Math.max(history.getLastRow() + 1, 4);
  const row = [label, daysGone, ...pcts, avgRating, maxStreak];
  history.getRange(lastRow, 1, 1, row.length).setValues([row])
         .setBackground('#262640').setFontColor('#ffffff');

  habits.getRange(2, 2, 15, 31).clearContent();
  habits.getRange(1, 2, 16, 31).setBorder(false, false, false, false, false, false);
  highlightToday();
  ui.alert(label + ' архивирован! Начат новый месяц.');
}
"""

if __name__ == '__main__':
    main()
