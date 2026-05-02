const APP_CONFIG = {
  timeZone: 'Asia/Tokyo',
  logSheetName: '勤怠ログ',
  summarySheetName: '月次集計',
  headers: [
    'ID',
    '種別',
    '日付',
    '開始',
    '終了',
    '内容',
    '分類',
    'コマ数',
    '分',
    '時間',
    'メモ',
    '作成日時',
    '更新日時',
  ],
  lessonType: 'レッスン',
  workType: '講師外業務',
  lessonCategories: ['通常コマ', 'キャンプ10時', 'その他・要確認'],
  lessonItemOptions: ['CB', 'CS', 'プラコ', '体験会', 'キャンプ', 'キャンプ(10時)', 'その他・要確認'],
  workCategories: ['教材開発', 'SNS', '研修', 'その他MTGなど'],
  workQuickItemsPropertyKey: 'workQuickItems',
  defaultWorkQuickItems: ['ULMTG', 'CBULMTG', 'チームMTG', 'UL業務', 'UL面談', '模擬レッスン'],
  maxWorkQuickItems: 6,
  lessonPayPerKoma: 2200,
  workHourlyPay: 1300,
};

const COL = {
  id: 1,
  type: 2,
  date: 3,
  start: 4,
  end: 5,
  content: 6,
  category: 7,
  koma: 8,
  minutes: 9,
  hours: 10,
  note: 11,
  createdAt: 12,
  updatedAt: 13,
};

function doGet() {
  setup();
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('勤怠・請求メモ')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('勤怠アプリ')
    .addItem('初期設定を実行', 'setup')
    .addToUi();
}

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone(APP_CONFIG.timeZone);

  const logSheet = ss.getSheetByName(APP_CONFIG.logSheetName) || ss.insertSheet(APP_CONFIG.logSheetName);
  prepareLogSheet_(logSheet);

  const summarySheet = ss.getSheetByName(APP_CONFIG.summarySheetName) || ss.insertSheet(APP_CONFIG.summarySheetName);
  prepareSummarySheet_(summarySheet);

  return {
    ok: true,
    message: '初期設定が完了しました。',
  };
}

function getInitialData() {
  setup();
  const now = new Date();
  return {
    todayIso: formatDate_(now, 'yyyy-MM-dd'),
    todayDisplay: formatDate_(now, 'yyyy年M月d日'),
    currentMonth: formatDate_(now, 'yyyy-MM'),
    lessonCategories: APP_CONFIG.lessonCategories,
    workCategories: APP_CONFIG.workCategories,
    lessonItemOptions: APP_CONFIG.lessonItemOptions,
    lessonQuickItems: APP_CONFIG.lessonItemOptions,
    workQuickItems: getWorkQuickItems_(),
    activeWork: findActiveWork_(),
    recentLogs: getRecentLogs_(5),
    salarySummary: getCurrentSalarySummary_(),
  };
}

function saveWorkQuickItems(items) {
  const normalized = normalizeWorkQuickItems_(items);
  PropertiesService.getUserProperties().setProperty(
    APP_CONFIG.workQuickItemsPropertyKey,
    JSON.stringify(normalized)
  );
  return {
    ok: true,
    message: 'よく使う内容を保存しました。',
    workQuickItems: normalized,
  };
}

function saveLesson(payload) {
  return withLock_(function () {
    setup();
    const date = parseDateInput_(payload.date) || dateOnly_(new Date());
    const lessonItems = normalizeLessonItems_(payload.lessonItems);
    const category = getLessonCategoryByItems_(lessonItems);
    const koma = Number(payload.koma);
    const note = trim_(payload.note);
    const content = lessonItems.join(' / ') + (note ? ' / ' + note : '');

    if (!koma || koma < 1 || koma > 4) {
      throw new Error('コマ数は1から4の範囲で選んでください。');
    }
    if (lessonItems.length !== koma) {
      throw new Error('コマ数とコマ内容の数が一致していません。');
    }

    const now = new Date();
    const row = [
      createId_(),
      APP_CONFIG.lessonType,
      date,
      '',
      '',
      content,
      category,
      koma,
      '',
      '',
      note,
      now,
      now,
    ];
    getLogSheet_().appendRow(row);
    applySheetFormats_();

    return {
      ok: true,
      message: 'レッスンを保存しました。',
      activeWork: findActiveWork_(),
      recentLogs: getRecentLogs_(5),
      salarySummary: getCurrentSalarySummary_(),
    };
  });
}

function startWork(payload) {
  return withLock_(function () {
    setup();
    const active = findActiveWork_();
    if (active) {
      throw new Error('未終了の講師外業務があります。先に終了してください。');
    }

    payload = payload || {};
    const content = trim_(payload.content);
    const category = payload.category
      ? assertInList_(payload.category, APP_CONFIG.workCategories, '講師外業務分類')
      : 'その他MTGなど';
    const note = trim_(payload.note);

    if (!content) {
      throw new Error('出勤時に内容を入力してください。');
    }

    const now = new Date();
    const row = [
      createId_(),
      APP_CONFIG.workType,
      dateOnly_(now),
      formatDate_(now, 'HH:mm'),
      '',
      content,
      category,
      '',
      '',
      '',
      note,
      now,
      now,
    ];
    getLogSheet_().appendRow(row);
    applySheetFormats_();

    return {
      ok: true,
      message: '出勤しました。',
      activeWork: findActiveWork_(),
      recentLogs: getRecentLogs_(5),
      salarySummary: getCurrentSalarySummary_(),
    };
  });
}

function prepareFinishWork() {
  setup();
  const active = findActiveWork_();
  if (!active) {
    throw new Error('開始中の講師外業務がありません。');
  }

  const now = new Date();
  return {
    endIso: formatDate_(now, "yyyy-MM-dd'T'HH:mm:ss"),
    endDisplay: formatDate_(now, 'yyyy/MM/dd HH:mm'),
  };
}

function finishWork(payload) {
  return withLock_(function () {
    setup();
    const activeRow = findActiveWorkRow_();
    if (!activeRow) {
      throw new Error('開始中の講師外業務がありません。');
    }

    payload = payload || {};
    const endAt = payload.endIso ? parseLocalDateTime_(payload.endIso) : new Date();
    const startAt = buildDateTime_(activeRow.values[COL.date - 1], activeRow.values[COL.start - 1]);

    if (!startAt) {
      throw new Error('開始時刻を読み取れませんでした。ログを確認してください。');
    }
    if (endAt.getTime() < startAt.getTime()) {
      throw new Error('終了時刻が開始時刻より前になっています。');
    }

    const minutes = Math.round((endAt.getTime() - startAt.getTime()) / 60000);
    const hours = roundHours_(minutes);
    const sheet = getLogSheet_();
    const row = activeRow.rowNumber;

    sheet.getRange(row, COL.end).setValue(formatDate_(endAt, 'HH:mm'));
    if (payload.content) {
      sheet.getRange(row, COL.content).setValue(trim_(payload.content));
    }
    if (payload.category) {
      sheet.getRange(row, COL.category).setValue(assertInList_(payload.category, APP_CONFIG.workCategories, '講師外業務分類'));
    }
    if (payload.note) {
      sheet.getRange(row, COL.note).setValue(trim_(payload.note));
    }
    sheet.getRange(row, COL.minutes).setValue(minutes);
    sheet.getRange(row, COL.hours).setValue(hours);
    sheet.getRange(row, COL.updatedAt).setValue(new Date());
    applySheetFormats_();

    return {
      ok: true,
      message: '退勤しました。',
      activeWork: findActiveWork_(),
      recentLogs: getRecentLogs_(5),
      salarySummary: getCurrentSalarySummary_(),
    };
  });
}

function cancelActiveWork() {
  return withLock_(function () {
    setup();
    const activeRow = findActiveWorkRow_();
    if (!activeRow) {
      throw new Error('取り消せる未退勤の出勤がありません。');
    }

    getLogSheet_().deleteRow(activeRow.rowNumber);
    return {
      ok: true,
      message: '出勤を取り消しました。',
      activeWork: findActiveWork_(),
      recentLogs: getRecentLogs_(5),
      salarySummary: getCurrentSalarySummary_(),
    };
  });
}

function deleteLog(id) {
  return withLock_(function () {
    setup();
    const found = findRowById_(id);
    if (!found) {
      throw new Error('削除対象のログが見つかりません。');
    }
    getLogSheet_().deleteRow(found.rowNumber);
    return {
      ok: true,
      message: 'ログを削除しました。',
      activeWork: findActiveWork_(),
      recentLogs: getRecentLogs_(5),
      salarySummary: getCurrentSalarySummary_(),
    };
  });
}

function updateLog(payload) {
  return withLock_(function () {
    setup();
    const found = findRowById_(payload.id);
    if (!found) {
      throw new Error('編集対象のログが見つかりません。');
    }

    const type = found.values[COL.type - 1];
    const date = parseDateInput_(payload.date);
    const content = trim_(payload.content);
    const note = trim_(payload.note);

    if (!date) {
      throw new Error('日付を入力してください。');
    }

    const sheet = getLogSheet_();
    const row = found.rowNumber;

    if (type === APP_CONFIG.lessonType) {
      const category = assertInList_(payload.category, APP_CONFIG.lessonCategories, 'レッスン分類');
      const koma = Number(payload.koma);
      if (!koma || koma < 1 || koma > 5) {
        throw new Error('コマ数は1から5の範囲で選んでください。');
      }

      sheet.getRange(row, COL.date, 1, 9).setValues([[
        date,
        '',
        '',
        content,
        category,
        koma,
        '',
        '',
        note,
      ]]);
    } else if (type === APP_CONFIG.workType) {
      const start = normalizeTimeText_(payload.start);
      const end = normalizeTimeText_(payload.end);
      const category = payload.category ? assertInList_(payload.category, APP_CONFIG.workCategories, '講師外業務分類') : '';
      if (!start) {
        throw new Error('開始時刻を入力してください。');
      }
      if (end && (!content || !category)) {
        throw new Error('終了済みの講師外業務は内容と分類を入力してください。');
      }

      let minutes = '';
      let hours = '';
      if (end) {
        const startAt = buildDateTime_(date, start);
        const endAt = buildDateTime_(date, end);
        if (endAt.getTime() < startAt.getTime()) {
          throw new Error('編集画面では日付をまたぐ時間は計算できません。必要な場合はシートで直接修正してください。');
        }
        minutes = Math.round((endAt.getTime() - startAt.getTime()) / 60000);
        hours = roundHours_(minutes);
      }

      sheet.getRange(row, COL.date, 1, 9).setValues([[
        date,
        start,
        end,
        content,
        category,
        '',
        minutes,
        hours,
        note,
      ]]);
    } else {
      throw new Error('不明な種別のログです。');
    }

    sheet.getRange(row, COL.updatedAt).setValue(new Date());
    applySheetFormats_();

    return {
      ok: true,
      message: 'ログを更新しました。',
      activeWork: findActiveWork_(),
      recentLogs: getRecentLogs_(5),
      salarySummary: getCurrentSalarySummary_(),
    };
  });
}

function getMonthlySummary(month) {
  setup();
  const target = parseMonth_(month);
  const logs = readLogObjects_();
  const summary = createEmptySummary_(target);

  logs.forEach(function (log) {
    if (log.month !== target.monthValue) {
      return;
    }

    if (log.type === APP_CONFIG.lessonType) {
      aggregateLesson_(summary, log);
      return;
    }

    if (log.type === APP_CONFIG.workType) {
      aggregateWork_(summary, log);
    }
  });

  finalizeSummary_(summary);
  writeMonthlySummary_(summary);
  return summary;
}

function getCurrentSalarySummary_() {
  return buildSalarySummary_(formatDate_(new Date(), 'yyyy-MM'));
}

function buildSalarySummary_(month) {
  const target = parseMonth_(month);
  const summary = {
    month: target.monthValue,
    monthLabel: target.monthLabel,
    lessonKoma: 0,
    workMinutes: 0,
    workHours: '0.00',
    lessonPay: 0,
    workPay: 0,
    totalPay: 0,
    lessonPayPerKoma: APP_CONFIG.lessonPayPerKoma,
    workHourlyPay: APP_CONFIG.workHourlyPay,
  };

  readLogObjects_().forEach(function (log) {
    if (log.month !== target.monthValue) {
      return;
    }

    if (log.type === APP_CONFIG.lessonType) {
      summary.lessonKoma += Number(log.koma || 0);
      return;
    }

    if (log.type === APP_CONFIG.workType && log.end) {
      let minutes = Number(log.minutes || 0);
      if (!minutes && log.start && log.end) {
        const startAt = buildDateTime_(log.dateIso, log.start);
        const endAt = buildDateTime_(log.dateIso, log.end);
        if (startAt && endAt && endAt.getTime() >= startAt.getTime()) {
          minutes = Math.round((endAt.getTime() - startAt.getTime()) / 60000);
        }
      }
      summary.workMinutes += minutes;
    }
  });

  summary.workHours = roundHours_(summary.workMinutes).toFixed(2);
  summary.lessonPay = summary.lessonKoma * APP_CONFIG.lessonPayPerKoma;
  summary.workPay = Math.round(Number(summary.workHours) * APP_CONFIG.workHourlyPay);
  summary.totalPay = summary.lessonPay + summary.workPay;
  return summary;
}

function prepareLogSheet_(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(APP_CONFIG.headers);
  } else {
    sheet.getRange(1, 1, 1, APP_CONFIG.headers.length).setValues([APP_CONFIG.headers]);
  }
  sheet.setFrozenRows(1);
  applySheetFormats_();
}

function prepareSummarySheet_(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, 2).setValues([['項目', '値']]);
  }
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, 8);
}

function getWorkQuickItems_() {
  const stored = PropertiesService.getUserProperties().getProperty(APP_CONFIG.workQuickItemsPropertyKey);
  if (!stored) {
    return APP_CONFIG.defaultWorkQuickItems.slice();
  }

  try {
    const parsed = JSON.parse(stored);
    return normalizeWorkQuickItems_(parsed);
  } catch (error) {
    return APP_CONFIG.defaultWorkQuickItems.slice();
  }
}

function normalizeWorkQuickItems_(items) {
  if (!Array.isArray(items)) {
    throw new Error('よく使う内容の形式が正しくありません。');
  }

  const seen = {};
  const normalized = [];
  items.forEach(function (item) {
    const text = trim_(item);
    if (!text || seen[text]) {
      return;
    }
    seen[text] = true;
    normalized.push(text.slice(0, 40));
  });

  if (normalized.length > APP_CONFIG.maxWorkQuickItems) {
    throw new Error('よく使う内容は最大' + APP_CONFIG.maxWorkQuickItems + '個までです。');
  }

  return normalized;
}

function normalizeLessonItems_(items) {
  if (!Array.isArray(items)) {
    throw new Error('コマ内容を選択してください。');
  }

  return items.map(function (item) {
    return assertInList_(item, APP_CONFIG.lessonItemOptions, 'コマ内容');
  });
}

function getLessonCategoryByItems_(items) {
  if (items.some(function (item) {
    return item === 'その他・要確認';
  })) {
    return 'その他・要確認';
  }

  const hasCamp10 = items.some(function (item) {
    return item === 'キャンプ(10時)';
  });
  const hasNormal = items.some(function (item) {
    return item !== 'キャンプ(10時)';
  });

  if (hasCamp10 && hasNormal) {
    return 'その他・要確認';
  }
  if (hasCamp10) {
    return 'キャンプ10時';
  }
  return '通常コマ';
}

function getLogSheet_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(APP_CONFIG.logSheetName);
}

function getSummarySheet_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(APP_CONFIG.summarySheetName);
}

function applySheetFormats_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(APP_CONFIG.logSheetName);
  if (!sheet) {
    return;
  }
  const maxRows = Math.max(sheet.getMaxRows() - 1, 1);
  sheet.getRange(2, COL.date, maxRows, 1).setNumberFormat('yyyy/mm/dd');
  sheet.getRange(2, COL.minutes, maxRows, 1).setNumberFormat('0');
  sheet.getRange(2, COL.hours, maxRows, 1).setNumberFormat('0.00');
  sheet.getRange(2, COL.createdAt, maxRows, 2).setNumberFormat('yyyy/mm/dd hh:mm');
  sheet.autoResizeColumns(1, APP_CONFIG.headers.length);
}

function readLogObjects_() {
  const sheet = getLogSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }

  return sheet
    .getRange(2, 1, lastRow - 1, APP_CONFIG.headers.length)
    .getValues()
    .map(function (values, index) {
      return rowToLog_(values, index + 2);
    })
    .filter(function (log) {
      return Boolean(log.id);
    });
}

function rowToLog_(values, rowNumber) {
  const dateValue = values[COL.date - 1];
  const createdAt = values[COL.createdAt - 1];
  const updatedAt = values[COL.updatedAt - 1];
  const minutes = toNumberOrBlank_(values[COL.minutes - 1]);
  const hours = toNumberOrBlank_(values[COL.hours - 1]);

  return {
    rowNumber: rowNumber,
    id: String(values[COL.id - 1] || ''),
    type: String(values[COL.type - 1] || ''),
    dateIso: formatValueDate_(dateValue, 'yyyy-MM-dd'),
    dateDisplay: formatValueDate_(dateValue, 'yyyy/MM/dd'),
    month: formatValueDate_(dateValue, 'yyyy-MM'),
    start: formatTimeValue_(values[COL.start - 1]),
    end: formatTimeValue_(values[COL.end - 1]),
    content: String(values[COL.content - 1] || ''),
    category: String(values[COL.category - 1] || ''),
    koma: toNumberOrBlank_(values[COL.koma - 1]),
    minutes: minutes,
    hours: hours === '' ? '' : Number(hours).toFixed(2),
    note: String(values[COL.note - 1] || ''),
    createdDisplay: formatValueDate_(createdAt, 'yyyy/MM/dd HH:mm'),
    updatedDisplay: formatValueDate_(updatedAt, 'yyyy/MM/dd HH:mm'),
    sortTime: Math.max(dateValueToMillis_(updatedAt), dateValueToMillis_(createdAt), rowNumber),
    isActive: values[COL.type - 1] === APP_CONFIG.workType && Boolean(values[COL.start - 1]) && !values[COL.end - 1],
  };
}

function getRecentLogs_(limit) {
  return readLogObjects_()
    .sort(function (a, b) {
      return (b.sortTime - a.sortTime) || (b.rowNumber - a.rowNumber);
    })
    .slice(0, limit || 5);
}

function findActiveWork_() {
  const activeRow = findActiveWorkRow_();
  return activeRow ? rowToLog_(activeRow.values, activeRow.rowNumber) : null;
}

function findActiveWorkRow_() {
  const sheet = getLogSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return null;
  }

  const values = sheet.getRange(2, 1, lastRow - 1, APP_CONFIG.headers.length).getValues();
  for (let i = 0; i < values.length; i += 1) {
    const row = values[i];
    if (row[COL.type - 1] === APP_CONFIG.workType && row[COL.start - 1] && !row[COL.end - 1]) {
      return {
        rowNumber: i + 2,
        values: row,
      };
    }
  }
  return null;
}

function findRowById_(id) {
  const sheet = getLogSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return null;
  }

  const values = sheet.getRange(2, 1, lastRow - 1, APP_CONFIG.headers.length).getValues();
  for (let i = 0; i < values.length; i += 1) {
    if (String(values[i][COL.id - 1]) === String(id)) {
      return {
        rowNumber: i + 2,
        values: values[i],
      };
    }
  }
  return null;
}

function createEmptySummary_(target) {
  return {
    month: target.monthValue,
    monthLabel: target.monthLabel,
    normalKoma: 0,
    camp10Koma: 0,
    otherConfirmKoma: 0,
    lessonTotalKoma: 0,
    workMinutes: {
      '教材開発': 0,
      SNS: 0,
      '研修': 0,
      'その他MTGなど': 0,
    },
    workHours: {
      '教材開発': '0.00',
      SNS: '0.00',
      '研修': '0.00',
      'その他MTGなど': '0.00',
    },
    workTotalMinutes: 0,
    workTotalHours: '0.00',
    unclassifiedCount: 0,
    unfinishedCount: 0,
    otherConfirmCount: 0,
    details: [],
    confirmNotes: [],
    warnings: [],
  };
}

function aggregateLesson_(summary, log) {
  const koma = Number(log.koma || 0);
  if (!APP_CONFIG.lessonCategories.includes(log.category)) {
    summary.unclassifiedCount += 1;
    summary.confirmNotes.push('未分類: ' + log.dateDisplay + ' レッスン ' + (log.content || '内容なし'));
    return;
  }

  if (log.category === '通常コマ') {
    summary.normalKoma += koma;
  } else if (log.category === 'キャンプ10時') {
    summary.camp10Koma += koma;
  } else if (log.category === 'その他・要確認') {
    summary.otherConfirmKoma += koma;
    summary.otherConfirmCount += 1;
    summary.confirmNotes.push('その他・要確認: ' + log.dateDisplay + ' ' + (log.content || '内容なし') + noteSuffix_(log.note));
  }
}

function aggregateWork_(summary, log) {
  if (!log.end) {
    summary.unfinishedCount += 1;
    summary.confirmNotes.push('終了漏れ: ' + log.dateDisplay + ' ' + (log.start || '--:--') + '開始 ' + (log.content || '内容未入力'));
    return;
  }

  if (!APP_CONFIG.workCategories.includes(log.category)) {
    summary.unclassifiedCount += 1;
    summary.confirmNotes.push('未分類: ' + log.dateDisplay + ' 講師外業務 ' + (log.content || '内容なし'));
    return;
  }

  let minutes = Number(log.minutes || 0);
  if (!minutes && log.start && log.end) {
    const startAt = buildDateTime_(log.dateIso, log.start);
    const endAt = buildDateTime_(log.dateIso, log.end);
    if (startAt && endAt && endAt.getTime() >= startAt.getTime()) {
      minutes = Math.round((endAt.getTime() - startAt.getTime()) / 60000);
    }
  }
  summary.workMinutes[log.category] += minutes;
  summary.details.push({
    date: log.dateDisplay,
    shortDate: shortDate_(log.dateIso),
    start: log.start,
    end: log.end,
    content: log.content,
    category: log.category,
    minutes: minutes,
    hours: roundHours_(minutes).toFixed(2),
    note: log.note,
  });
}

function finalizeSummary_(summary) {
  summary.lessonTotalKoma = summary.normalKoma + summary.camp10Koma;

  APP_CONFIG.workCategories.forEach(function (category) {
    const minutes = summary.workMinutes[category] || 0;
    summary.workTotalMinutes += minutes;
    summary.workHours[category] = roundHours_(minutes).toFixed(2);
  });
  summary.workTotalHours = roundHours_(summary.workTotalMinutes).toFixed(2);

  if (summary.otherConfirmCount > 0) {
    summary.warnings.push('その他・要確認のレッスンがあります。内容を確認してください。');
  }
  if (summary.unfinishedCount > 0) {
    summary.warnings.push('終了漏れの講師外業務があります。勤怠ログを確認してください。');
  }
  if (summary.unclassifiedCount > 0) {
    summary.warnings.push('未分類のログがあります。分類を確認してください。');
  }
}

function writeMonthlySummary_(summary) {
  const sheet = getSummarySheet_();
  sheet.clearContents();

  const rows = [
    ['項目', '値'],
    ['対象月', summary.monthLabel],
    ['通常コマ数', summary.normalKoma],
    ['キャンプ10時コマ数', summary.camp10Koma],
    ['レッスン合計コマ数', summary.lessonTotalKoma],
    ['その他・要確認コマ数', summary.otherConfirmKoma],
    ['教材開発時間', summary.workHours['教材開発']],
    ['SNS時間', summary.workHours.SNS],
    ['研修時間', summary.workHours['研修']],
    ['その他MTGなど時間', summary.workHours['その他MTGなど']],
    ['講師外業務合計時間', summary.workTotalHours],
    ['未分類件数', summary.unclassifiedCount],
    ['終了漏れ件数', summary.unfinishedCount],
    ['その他・要確認件数', summary.otherConfirmCount],
    [],
    ['講師外業務の詳細一覧'],
    ['日付', '開始', '終了', '内容', '分類', '分', '時間', 'メモ'],
  ];

  summary.details.forEach(function (detail) {
    rows.push([
      detail.date,
      detail.start,
      detail.end,
      detail.content,
      detail.category,
      detail.minutes,
      detail.hours,
      detail.note,
    ]);
  });

  if (summary.confirmNotes.length > 0) {
    rows.push([]);
    rows.push(['要確認メモ']);
    summary.confirmNotes.forEach(function (note) {
      rows.push([note]);
    });
  }

  const width = Math.max.apply(null, rows.map(function (row) {
    return row.length;
  }));
  const padded = rows.map(function (row) {
    const copy = row.slice();
    while (copy.length < width) {
      copy.push('');
    }
    return copy;
  });

  sheet.getRange(1, 1, padded.length, width).setValues(padded);
  sheet.getRange(1, 1, 1, width).setFontWeight('bold');
  sheet.autoResizeColumns(1, width);
}

function parseMonth_(month) {
  const match = String(month || '').match(/^(\d{4})-(\d{2})$/);
  if (!match) {
    throw new Error('対象月を選択してください。');
  }
  const year = Number(match[1]);
  const monthNumber = Number(match[2]);
  if (monthNumber < 1 || monthNumber > 12) {
    throw new Error('対象月が正しくありません。');
  }
  return {
    year: year,
    monthNumber: monthNumber,
    monthValue: match[1] + '-' + match[2],
    monthLabel: year + '年' + monthNumber + '月',
  };
}

function parseDateInput_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return dateOnly_(value);
  }

  const match = String(value || '').match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/);
  if (!match) {
    return null;
  }
  return new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
}

function parseLocalDateTime_(value) {
  const match = String(value || '').match(/^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})(?::(\d{2}))?$/);
  if (!match) {
    throw new Error('終了時刻を読み取れませんでした。');
  }
  return new Date(
    Number(match[1]),
    Number(match[2]) - 1,
    Number(match[3]),
    Number(match[4]),
    Number(match[5]),
    Number(match[6] || 0)
  );
}

function buildDateTime_(dateValue, timeValue) {
  const date = parseDateInput_(formatValueDate_(dateValue, 'yyyy-MM-dd'));
  const time = normalizeTimeText_(timeValue);
  if (!date || !time) {
    return null;
  }
  const parts = time.split(':').map(Number);
  return new Date(date.getFullYear(), date.getMonth(), date.getDate(), parts[0], parts[1], 0, 0);
}

function normalizeTimeText_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return formatDate_(value, 'HH:mm');
  }

  const match = String(value || '').trim().match(/^(\d{1,2}):(\d{2})$/);
  if (!match) {
    return '';
  }
  const hour = Number(match[1]);
  const minute = Number(match[2]);
  if (hour < 0 || hour > 23 || minute < 0 || minute > 59) {
    return '';
  }
  return String(hour).padStart(2, '0') + ':' + String(minute).padStart(2, '0');
}

function assertInList_(value, list, label) {
  const text = trim_(value);
  if (!list.includes(text)) {
    throw new Error(label + 'を選択してください。');
  }
  return text;
}

function withLock_(callback) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(10000);
  try {
    return callback();
  } finally {
    lock.releaseLock();
  }
}

function createId_() {
  return formatDate_(new Date(), 'yyyyMMddHHmmss') + '-' + Math.floor(Math.random() * 100000);
}

function dateOnly_(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

function formatDate_(date, pattern) {
  return Utilities.formatDate(date, APP_CONFIG.timeZone, pattern);
}

function formatValueDate_(value, pattern) {
  if (!value) {
    return '';
  }
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return formatDate_(value, pattern);
  }
  const parsed = parseDateInput_(value);
  return parsed ? formatDate_(parsed, pattern) : String(value);
}

function formatTimeValue_(value) {
  if (!value) {
    return '';
  }
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return formatDate_(value, 'HH:mm');
  }
  return normalizeTimeText_(value) || String(value);
}

function shortDate_(dateIso) {
  const match = String(dateIso || '').match(/^\d{4}-(\d{2})-(\d{2})$/);
  if (!match) {
    return '';
  }
  return Number(match[1]) + '/' + Number(match[2]);
}

function toNumberOrBlank_(value) {
  if (value === '' || value === null || typeof value === 'undefined') {
    return '';
  }
  const number = Number(value);
  return isNaN(number) ? '' : number;
}

function dateValueToMillis_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return value.getTime();
  }
  const parsed = new Date(String(value || ''));
  return isNaN(parsed.getTime()) ? 0 : parsed.getTime();
}

function roundHours_(minutes) {
  return Math.round((Number(minutes || 0) / 60) * 100) / 100;
}

function trim_(value) {
  return String(value || '').trim();
}

function noteSuffix_(note) {
  return note ? ' / メモ: ' + note : '';
}
