
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
