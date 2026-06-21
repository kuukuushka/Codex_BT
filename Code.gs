const SHEET_NAME = 'Ежедневный_отчет';
const HEADERS = ['ID', 'Дата', 'Время', 'Сотрудник', 'Тип действия', 'Данные'];

const EMPLOYEES = [
  'Сотрудник 1',
  'Сотрудник 2',
  'Сотрудник 3'
];

const ACTION_TYPES = [
  'Связи',
  'Кошельки',
  'Тикеты',
  'Потенциальные агенты',
  'Проверка кошельков',
  'Заблокированные устройства',
  'Заблокированные кошельки',
  'Заблокированные игроки',
  'Запрос транзакций',
  'Авторезерв'
];

function setup() {
  getOrCreateSheet();
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Ежедневный отчет')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getInitialData() {
  setup();

  return {
    employees: EMPLOYEES,
    actionTypes: ACTION_TYPES,
    entries: getTodayEntries()
  };
}

function addEntry(payload) {
  const employee = String(payload && payload.employee || '').trim();
  const actionType = String(payload && payload.actionType || '').trim();
  const data = String(payload && payload.data || '').trim();

  if (!employee) {
    throw new Error('Выберите сотрудника.');
  }
  if (EMPLOYEES.indexOf(employee) === -1) {
    throw new Error('Выбран неизвестный сотрудник.');
  }
  if (!actionType) {
    throw new Error('Не указан тип действия.');
  }
  if (ACTION_TYPES.indexOf(actionType) === -1) {
    throw new Error('Выбран неизвестный тип действия.');
  }
  if (!data) {
    throw new Error('Введите данные.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const sheet = getOrCreateSheet();
    const now = new Date();
    const timeZone = Session.getScriptTimeZone();
    const entry = {
      id: Utilities.getUuid(),
      date: Utilities.formatDate(now, timeZone, 'dd.MM.yyyy'),
      time: Utilities.formatDate(now, timeZone, 'HH:mm:ss'),
      employee: employee,
      actionType: actionType,
      data: data
    };

    sheet.appendRow([
      entry.id,
      entry.date,
      entry.time,
      entry.employee,
      entry.actionType,
      entry.data
    ]);

    return entry;
  } finally {
    lock.releaseLock();
  }
}

function deleteEntry(id) {
  const entryId = String(id || '').trim();

  if (!entryId) {
    throw new Error('Нельзя удалить запись без ID.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const sheet = getOrCreateSheet();
    const lastRow = sheet.getLastRow();

    if (lastRow < 2) {
      throw new Error('Запись не найдена.');
    }

    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();

    for (let index = 0; index < ids.length; index++) {
      if (String(ids[index][0]) === entryId) {
        sheet.deleteRow(index + 2);
        return getTodayEntries();
      }
    }

    throw new Error('Запись не найдена.');
  } finally {
    lock.releaseLock();
  }
}

function getTodayEntries() {
  const sheet = getOrCreateSheet();
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return [];
  }

  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd.MM.yyyy');
  const values = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();

  return values
    .filter(function(row) {
      return String(row[1]) === today;
    })
    .map(function(row) {
      return {
        id: String(row[0]),
        date: String(row[1]),
        time: String(row[2]),
        employee: String(row[3]),
        actionType: String(row[4]),
        data: String(row[5])
      };
    });
}

function getOrCreateSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  if (!spreadsheet) {
    throw new Error('Проект должен быть привязан к Google Таблице.');
  }

  let sheet = spreadsheet.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_NAME);
  }

  const currentHeaders = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  const shouldWriteHeaders = HEADERS.some(function(header, index) {
    return currentHeaders[index] !== header;
  });

  if (shouldWriteHeaders) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  }

  sheet.setFrozenRows(1);
  sheet.hideColumns(1);
  sheet.autoResizeColumns(2, HEADERS.length - 1);

  return sheet;
}
