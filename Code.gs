// ==============================
// Life OS — GAS版
// ==============================

// ==============================
// 定数
// ==============================
var SHEET_DEFINITIONS = {
  diary:    ['id', 'datetime', 'mood', 'tag', 'text'],
  todo:     ['id', 'datetime', 'text', 'tag', 'priority', 'due', 'status', 'done_at'],
  books:    ['id', 'datetime', 'isbn', 'title', 'author', 'cover_url', 'medium', 'rating', 'status', 'note'],
  settings: ['key', 'value']
};

var SETTINGS_DEFAULTS = [
  ['gemini_api_key', ''],
  ['character_name', 'ピアちゃん'],
  ['character_prompt', 'あなたは「ピアちゃん」というキャラクターです。\n見た目はもちもちしたピンクのゆるキャラ。\n性格はのんびりしているけど、実はよく見ている。\n口調は「〜だよ」「〜だね」「〜かも！」。\n褒めるときは「すごいじゃん！」「がんばったね〜！」としっかり褒める。\n気になる点は「ちょっと気になったんだけど〜」とやさしく切り出す。\n絶対に説教しない。短めに話す。絵文字を適度に使う。\nユーザーの日記・ToDo・読書データにアクセスできる。\nデータに基づいた具体的なアドバイスをする。\n回答は簡潔に、200文字以内を目安にしてください。'],
  ['report_email', ''],
  ['spreadsheet_url', '']
];

// ==============================
// Web App ルーティング
// ==============================
function doGet(e) {
  var page = (e && e.parameter && e.parameter.p || '').toLowerCase();
  var file = page === 'books' ? 'Books' : page === 'chat' ? 'Chat' : 'UI';
  return HtmlService.createHtmlOutputFromFile(file)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setTitle('Life OS');
}

// ==============================
// スプレッドシート ヘルパー（プライベート）
// ==============================
function getOrCreateSheet_(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    // settingsシートは初期値を挿入
    if (name === 'settings') {
      SETTINGS_DEFAULTS.forEach(function(row) {
        sheet.appendRow(row);
      });
      // spreadsheet_url を自動設定
      var urlRow = findRowIndex_(sheet, 0, 'spreadsheet_url');
      if (urlRow > 0) {
        sheet.getRange(urlRow, 2).setValue(ss.getUrl());
      }
    }
  }
  return sheet;
}

function getSheetData_(sheetName) {
  var headers = SHEET_DEFINITIONS[sheetName];
  var sheet = getOrCreateSheet_(sheetName, headers);
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  var headerRow = data[0];
  var results = [];
  for (var i = 1; i < data.length; i++) {
    var obj = {};
    for (var j = 0; j < headerRow.length; j++) {
      var val = data[i][j];
      obj[headerRow[j]] = (val === '' || val === null || val === undefined) ? null : val;
    }
    results.push(obj);
  }
  return results;
}

function findRowIndex_(sheet, colIndex, value) {
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][colIndex]) === String(value)) return i + 1; // 1-based
  }
  return -1;
}

function generateId_() {
  return Utilities.getUuid();
}

function jstNow_() {
  var now = new Date();
  var jst = new Date(now.getTime() + 9 * 60 * 60 * 1000);
  return jst.toISOString().replace('Z', '+09:00');
}

function jstToday_() {
  var now = new Date();
  var jst = new Date(now.getTime() + 9 * 60 * 60 * 1000);
  return jst.toISOString().slice(0, 10);
}

// ==============================
// 設定
// ==============================
function getSettings() {
  var data = getSheetData_('settings');
  var map = {};
  data.forEach(function(row) {
    if (row.key) map[row.key] = row.value || '';
  });
  return map;
}

function saveSetting(key, value) {
  var sheet = getOrCreateSheet_('settings', SHEET_DEFINITIONS.settings);
  var rowIndex = findRowIndex_(sheet, 0, key);
  if (rowIndex > 0) {
    sheet.getRange(rowIndex, 2).setValue(value);
  } else {
    sheet.appendRow([key, value]);
  }
  return { key: key, value: value };
}

// ==============================
// 日記機能
// ==============================
function saveEntry(data) {
  var sheet = getOrCreateSheet_('diary', SHEET_DEFINITIONS.diary);
  var id = generateId_();
  var datetime = data.datetime || jstNow_();
  sheet.appendRow([id, datetime, data.mood || null, data.tag || null, data.text || null]);
  return { id: id, datetime: datetime };
}

function updateEntry(data) {
  var sheet = getOrCreateSheet_('diary', SHEET_DEFINITIONS.diary);
  var rowIndex = findRowIndex_(sheet, 0, data.id);
  if (rowIndex < 0) throw new Error('エントリが見つかりません');
  if (data.datetime !== undefined) sheet.getRange(rowIndex, 2).setValue(data.datetime);
  if (data.mood !== undefined) sheet.getRange(rowIndex, 3).setValue(data.mood);
  if (data.tag !== undefined) sheet.getRange(rowIndex, 4).setValue(data.tag);
  if (data.text !== undefined) sheet.getRange(rowIndex, 5).setValue(data.text);
  return { id: data.id, updated: true };
}

function deleteEntry(id) {
  var sheet = getOrCreateSheet_('diary', SHEET_DEFINITIONS.diary);
  var rowIndex = findRowIndex_(sheet, 0, id);
  if (rowIndex > 0) sheet.deleteRow(rowIndex);
  return { id: id, deleted: true };
}

// ==============================
// ToDo機能
// ==============================
function saveTodo(data) {
  var sheet = getOrCreateSheet_('todo', SHEET_DEFINITIONS.todo);
  var id = generateId_();
  var datetime = data.datetime || jstNow_();
  sheet.appendRow([id, datetime, data.text || '', data.tag || null, data.priority || 'mid', data.due || null, 'open', null]);
  return { id: id, datetime: datetime, status: 'open' };
}

function completeTodo(id) {
  var sheet = getOrCreateSheet_('todo', SHEET_DEFINITIONS.todo);
  var rowIndex = findRowIndex_(sheet, 0, id);
  if (rowIndex < 0) throw new Error('ToDoが見つかりません');
  sheet.getRange(rowIndex, 7).setValue('done'); // status
  sheet.getRange(rowIndex, 8).setValue(jstNow_()); // done_at
  return { id: id, status: 'done' };
}

function reopenTodo(id) {
  var sheet = getOrCreateSheet_('todo', SHEET_DEFINITIONS.todo);
  var rowIndex = findRowIndex_(sheet, 0, id);
  if (rowIndex < 0) throw new Error('ToDoが見つかりません');
  sheet.getRange(rowIndex, 7).setValue('open');
  sheet.getRange(rowIndex, 8).setValue('');
  return { id: id, status: 'open' };
}

function updateTodo(data) {
  var sheet = getOrCreateSheet_('todo', SHEET_DEFINITIONS.todo);
  var rowIndex = findRowIndex_(sheet, 0, data.id);
  if (rowIndex < 0) throw new Error('ToDoが見つかりません');
  // cols: id(1), datetime(2), text(3), tag(4), priority(5), due(6), status(7), done_at(8)
  if (data.text !== undefined) sheet.getRange(rowIndex, 3).setValue(data.text);
  if (data.tag !== undefined) sheet.getRange(rowIndex, 4).setValue(data.tag || '');
  if (data.priority !== undefined) sheet.getRange(rowIndex, 5).setValue(data.priority);
  if (data.due !== undefined) sheet.getRange(rowIndex, 6).setValue(data.due || '');
  if (data.status !== undefined) {
    sheet.getRange(rowIndex, 7).setValue(data.status);
    if (data.status === 'done') {
      sheet.getRange(rowIndex, 8).setValue(jstNow_());
    } else {
      sheet.getRange(rowIndex, 8).setValue('');
    }
  }
  return { id: data.id, updated: true };
}

function deleteTodo(id) {
  var sheet = getOrCreateSheet_('todo', SHEET_DEFINITIONS.todo);
  var rowIndex = findRowIndex_(sheet, 0, id);
  if (rowIndex > 0) sheet.deleteRow(rowIndex);
  return { id: id, deleted: true };
}

// ==============================
// タグサジェスト
// ==============================
function getTagSuggest(query) {
  var diaryData = getSheetData_('diary');
  var todoData = getSheetData_('todo');
  var freq = {};
  var allData = diaryData.concat(todoData);
  allData.forEach(function(row) {
    var tag = row.tag;
    if (tag && tag !== '') {
      var t = String(tag);
      if (!query || t.indexOf(query) >= 0) {
        freq[t] = (freq[t] || 0) + 1;
      }
    }
  });
  var tags = Object.keys(freq).map(function(tag) {
    return { tag: tag, count: freq[tag] };
  });
  tags.sort(function(a, b) { return b.count - a.count; });
  return { tags: tags.slice(0, 20) };
}

// ==============================
// ダッシュボード
// ==============================
function getDashboard() {
  var today = jstToday_();
  var diaryData = getSheetData_('diary');
  var todoData = getSheetData_('todo');

  // 今日のエントリ
  var todayEntries = diaryData.filter(function(e) {
    return e.datetime && String(e.datetime).indexOf(today) === 0;
  }).sort(function(a, b) {
    return String(b.datetime).localeCompare(String(a.datetime));
  });

  // 未完了ToDo（priority順→due順）
  var openTodos = todoData.filter(function(t) { return t.status === 'open'; });
  var priorityOrder = { high: 0, mid: 1, low: 2 };
  openTodos.sort(function(a, b) {
    var pa = priorityOrder[a.priority] !== undefined ? priorityOrder[a.priority] : 1;
    var pb = priorityOrder[b.priority] !== undefined ? priorityOrder[b.priority] : 1;
    if (pa !== pb) return pa - pb;
    // due null は後ろ
    if (a.due && !b.due) return -1;
    if (!a.due && b.due) return 1;
    if (a.due && b.due) return String(a.due).localeCompare(String(b.due));
    return 0;
  });
  openTodos = openTodos.slice(0, 30);

  // 過去同日の振り返り
  var monthDay = today.slice(5); // MM-DD
  var lookback = diaryData.filter(function(e) {
    var dt = String(e.datetime);
    return dt.indexOf(monthDay) >= 5 && dt.indexOf(today) !== 0;
  }).sort(function(a, b) {
    return String(b.datetime).localeCompare(String(a.datetime));
  }).slice(0, 5);

  // 最近の完了ToDo（5件）
  var doneTodos = todoData.filter(function(t) { return t.status === 'done'; });
  doneTodos.sort(function(a, b) {
    return String(b.done_at || '').localeCompare(String(a.done_at || ''));
  });
  var recentDone = doneTodos.slice(0, 5);

  // 30日ストリーク
  var streakMap = {};
  diaryData.forEach(function(e) {
    if (e.datetime) {
      var date = String(e.datetime).slice(0, 10);
      streakMap[date] = (streakMap[date] || 0) + 1;
    }
  });
  var streakData = [];
  for (var i = 29; i >= 0; i--) {
    var d = new Date(today + 'T00:00:00Z');
    d.setUTCDate(d.getUTCDate() - i);
    var key = d.toISOString().slice(0, 10);
    if (streakMap[key]) {
      streakData.push({ date: key, count: streakMap[key] });
    }
  }

  // サマリー統計
  var todayDoneCount = todoData.filter(function(t) {
    return t.status === 'done' && t.done_at && String(t.done_at).indexOf(today) === 0;
  }).length;

  var moodEntries = todayEntries.filter(function(e) { return e.mood && e.mood > 0; });
  var todayAvgMood = null;
  if (moodEntries.length > 0) {
    var sum = moodEntries.reduce(function(s, e) { return s + Number(e.mood); }, 0);
    todayAvgMood = Math.round((sum / moodEntries.length) * 10) / 10;
  }

  var overdueCount = openTodos.filter(function(t) { return t.due && String(t.due) < today; }).length;

  // ストリークカウント（連続記録日数）
  var streakCount = 0;
  for (var j = 0; j < 30; j++) {
    var sd = new Date(today + 'T00:00:00Z');
    sd.setUTCDate(sd.getUTCDate() - j);
    var sk = sd.toISOString().slice(0, 10);
    if (streakMap[sk]) {
      streakCount++;
    } else {
      break;
    }
  }

  // カレンダー予定
  var calendarEvents = [];
  try {
    var cal = CalendarApp.getDefaultCalendar();
    var todayDate = new Date(today + 'T00:00:00+09:00');
    var tomorrowDate = new Date(todayDate.getTime() + 86400000);
    var dayAfterDate = new Date(todayDate.getTime() + 2 * 86400000);

    var todayEvents = cal.getEvents(todayDate, tomorrowDate);
    var tomorrowEvents = cal.getEvents(tomorrowDate, dayAfterDate);

    todayEvents.forEach(function(ev) {
      calendarEvents.push({
        title: ev.getTitle(),
        date: today,
        startTime: ev.isAllDayEvent() ? '終日' : Utilities.formatDate(ev.getStartTime(), 'Asia/Tokyo', 'HH:mm'),
        endTime: ev.isAllDayEvent() ? '' : Utilities.formatDate(ev.getEndTime(), 'Asia/Tokyo', 'HH:mm'),
        allDay: ev.isAllDayEvent(),
        location: ev.getLocation() || ''
      });
    });

    var tomorrowStr = tomorrowDate.toISOString().slice(0, 10);
    tomorrowEvents.forEach(function(ev) {
      calendarEvents.push({
        title: ev.getTitle(),
        date: tomorrowStr,
        startTime: ev.isAllDayEvent() ? '終日' : Utilities.formatDate(ev.getStartTime(), 'Asia/Tokyo', 'HH:mm'),
        endTime: ev.isAllDayEvent() ? '' : Utilities.formatDate(ev.getEndTime(), 'Asia/Tokyo', 'HH:mm'),
        allDay: ev.isAllDayEvent(),
        location: ev.getLocation() || ''
      });
    });
  } catch (e) {
    // カレンダーアクセスエラーは無視
  }

  return {
    today: today,
    todayEntries: todayEntries,
    openTodos: openTodos,
    lookback: lookback,
    recentDone: recentDone,
    streakData: streakData,
    calendarEvents: calendarEvents,
    summary: {
      todayEntryCount: todayEntries.length,
      openCount: openTodos.length,
      todayDoneCount: todayDoneCount,
      overdueCount: overdueCount,
      todayAvgMood: todayAvgMood,
      streakCount: streakCount
    }
  };
}

// ==============================
// 読書機能
// ==============================
function searchBook(query) {
  if (!query) return { books: [] };
  var q = query.trim();
  var cleaned = q.replace(/[\-\s]/g, '');

  var isIsbn13 = /^(978|979)\d{10}$/.test(cleaned);
  var isIsbn10 = /^\d{9}[\dXx]$/.test(cleaned);
  var isIsbn = isIsbn13 || isIsbn10;
  var is13Digits = /^\d{13}$/.test(cleaned);
  var isNonIsbnBarcode = is13Digits && !isIsbn13;

  if (isNonIsbnBarcode) {
    return { books: [], hint: 'これはISBNバーコードではないみたい📖 上のバーコード（978で始まる方）を読み取ってね！' };
  }

  if (isIsbn) {
    return searchByIsbn_(cleaned);
  } else {
    return searchByTitle_(q);
  }
}

function searchByIsbn_(isbn) {
  try {
    var res = UrlFetchApp.fetch('https://api.openbd.jp/v1/get?isbn=' + isbn, { muteHttpExceptions: true });
    var data = JSON.parse(res.getContentText());
    if (data && data[0]) {
      var s = data[0].summary;
      return {
        books: [{
          isbn: s.isbn || isbn,
          title: s.title || '',
          author: s.author || '',
          cover_url: s.cover || null,
          publisher: s.publisher || '',
          published_date: s.pubdate || ''
        }]
      };
    }
  } catch (e) {}
  return searchByTitle_(isbn);
}

function searchByTitle_(q) {
  // 検索クエリ正規化
  var normalized = q.replace(/\u3000/g, ' ').replace(/[はがのをにでと]/g, ' ').replace(/\s+/g, ' ').trim();

  try {
    var ndlUrl = 'https://ndlsearch.ndl.go.jp/api/opensearch?title=' + encodeURIComponent(normalized) + '&cnt=10&dpid=iss-ndl-opac';
    var res = UrlFetchApp.fetch(ndlUrl, {
      muteHttpExceptions: true,
      headers: { 'User-Agent': 'LifeOS/1.0 (personal-app)' }
    });

    if (res.getResponseCode() !== 200) throw new Error('NDL HTTP ' + res.getResponseCode());

    var xmlText = res.getContentText();
    var books = parseNdlXml_(xmlText);

    if (books.length > 0) {
      return { books: books };
    }

    // NDLで見つからない場合はOpen Libraryにフォールバック
    return searchOpenLibrary_(q);
  } catch (e) {
    return searchOpenLibrary_(q);
  }
}

function parseNdlXml_(xmlText) {
  var books = [];
  try {
    // XMLパースの前に不正な文字を除去
    xmlText = xmlText.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '');
    var doc = XmlService.parse(xmlText);
    var root = doc.getRootElement();
    var ns = XmlService.getNamespace('http://purl.org/rss/1.0/');
    var dcNs = XmlService.getNamespace('dc', 'http://purl.org/dc/elements/1.1/');

    var channel = root.getChild('channel', ns);
    if (!channel) return books;

    var items = channel.getChildren('item', ns);
    if (!items || items.length === 0) {
      // RSSフォーマットではなくAtomかもしれないので、正規表現フォールバック
      return parseNdlRegex_(xmlText);
    }

    items.forEach(function(item) {
      var rawTitle = '';
      var titleEl = item.getChild('title', ns);
      if (titleEl) rawTitle = titleEl.getText();
      var title = rawTitle.split('/')[0].trim();
      if (!title) return;

      var author = '';
      var creatorEl = item.getChild('creator', dcNs);
      if (creatorEl) author = creatorEl.getText();

      var publisher = '';
      var pubEl = item.getChild('publisher', dcNs);
      if (pubEl) publisher = pubEl.getText();

      var published_date = '';
      var dateEl = item.getChild('date', dcNs);
      if (dateEl) published_date = dateEl.getText().slice(0, 4);

      // ISBN: 978/979で始まる13桁
      var isbn = null;
      var desc = item.getChild('description', ns);
      var fullText = (desc ? desc.getText() : '') + ' ' + xmlText.slice(xmlText.indexOf(title), xmlText.indexOf(title) + 500);
      var isbnMatch = fullText.match(/(?:978|979)\d{10}/);
      if (isbnMatch) isbn = isbnMatch[0];

      // identifierからISBN取得を試みる
      if (!isbn) {
        var identifiers = item.getChildren('identifier', dcNs);
        if (identifiers) {
          identifiers.forEach(function(idEl) {
            var idText = idEl.getText();
            var m = idText.match(/(?:978|979)\d{10}/);
            if (m) isbn = m[0];
          });
        }
      }

      books.push({
        isbn: isbn,
        title: title,
        author: author,
        publisher: publisher,
        published_date: published_date,
        cover_url: isbn ? 'https://books.google.com/books/content?vid=isbn:' + isbn + '&printsec=frontcover&img=1&zoom=1' : null
      });
    });
  } catch (e) {
    return parseNdlRegex_(xmlText);
  }
  return books;
}

function parseNdlRegex_(xml) {
  var books = [];
  var itemRegex = /<item>([\s\S]*?)<\/item>/g;
  var match;
  while ((match = itemRegex.exec(xml)) !== null) {
    var item = match[1];
    var rawTitle = tagContent_(item, 'title') || '';
    var title = rawTitle.split('/')[0].trim();
    if (!title) continue;

    var author = tagContent_(item, 'dc:creator') || '';
    var publisher = tagContent_(item, 'dc:publisher') || '';
    var published_date = (tagContent_(item, 'dc:date') || '').slice(0, 4);

    var isbn = null;
    var isbnMatch = item.match(/(?:978|979)\d{10}/);
    if (isbnMatch) isbn = isbnMatch[0];

    books.push({
      isbn: isbn,
      title: title,
      author: author,
      publisher: publisher,
      published_date: published_date,
      cover_url: isbn ? 'https://books.google.com/books/content?vid=isbn:' + isbn + '&printsec=frontcover&img=1&zoom=1' : null
    });
  }
  return books;
}

function tagContent_(xml, tag) {
  var m = xml.match(new RegExp('<' + tag + '[^>]*>([\\s\\S]*?)<\\/' + tag + '>', 'i'));
  return m ? m[1].trim() : null;
}

function searchOpenLibrary_(q) {
  try {
    var url = 'https://openlibrary.org/search.json?q=' + encodeURIComponent(q) + '&limit=10';
    var res = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      headers: { 'User-Agent': 'LifeOS/1.0 (personal-app)' }
    });
    var data = JSON.parse(res.getContentText());
    if (!data.docs || !data.docs.length) return { books: [] };

    var books = data.docs.map(function(doc) {
      var isbn = (doc.isbn || [])[0] || null;
      var coverId = doc.cover_i;
      return {
        isbn: isbn,
        title: doc.title || '',
        author: (doc.author_name || []).join(', '),
        cover_url: coverId ? 'https://covers.openlibrary.org/b/id/' + coverId + '-M.jpg' : null,
        publisher: (doc.publisher || [])[0] || '',
        published_date: doc.first_publish_year ? String(doc.first_publish_year) : ''
      };
    });
    return { books: books };
  } catch (e) {
    return { books: [] };
  }
}

function saveBook(data) {
  var sheet = getOrCreateSheet_('books', SHEET_DEFINITIONS.books);
  var id = generateId_();
  var datetime = jstNow_();
  sheet.appendRow([
    id, datetime,
    data.isbn || null, data.title || null, data.author || null,
    data.cover_url || null, data.medium || null, data.rating || null,
    data.status || 'done', data.note || null
  ]);
  return { id: id, datetime: datetime };
}

function updateBook(data) {
  var sheet = getOrCreateSheet_('books', SHEET_DEFINITIONS.books);
  var rowIndex = findRowIndex_(sheet, 0, data.id);
  if (rowIndex < 0) throw new Error('本が見つかりません');
  // cols: id(1), datetime(2), isbn(3), title(4), author(5), cover_url(6), medium(7), rating(8), status(9), note(10)
  if (data.medium !== undefined) sheet.getRange(rowIndex, 7).setValue(data.medium);
  if (data.rating !== undefined) sheet.getRange(rowIndex, 8).setValue(data.rating);
  if (data.status !== undefined) sheet.getRange(rowIndex, 9).setValue(data.status);
  if (data.note !== undefined) sheet.getRange(rowIndex, 10).setValue(data.note);
  return { id: data.id, updated: true };
}

function deleteBook(id) {
  var sheet = getOrCreateSheet_('books', SHEET_DEFINITIONS.books);
  var rowIndex = findRowIndex_(sheet, 0, id);
  if (rowIndex > 0) sheet.deleteRow(rowIndex);
  return { id: id, deleted: true };
}

function getRecentBooks(limit, statusFilter, mediumFilter, sort, search) {
  var data = getSheetData_('books');

  // ステータスフィルタ
  if (statusFilter && statusFilter !== 'all') {
    data = data.filter(function(b) { return b.status === statusFilter; });
  }

  // 媒体フィルタ
  if (mediumFilter && mediumFilter !== 'all') {
    data = data.filter(function(b) {
      return b.medium && String(b.medium).toLowerCase() === String(mediumFilter).toLowerCase();
    });
  }

  // 検索
  if (search) {
    var q = search.toLowerCase();
    data = data.filter(function(b) {
      return (b.title && String(b.title).toLowerCase().indexOf(q) >= 0) ||
             (b.author && String(b.author).toLowerCase().indexOf(q) >= 0) ||
             (b.note && String(b.note).toLowerCase().indexOf(q) >= 0);
    });
  }

  // ソート
  var sortKey = sort || 'datetime_desc';
  if (sortKey === 'datetime_asc') {
    data.sort(function(a, b) { return String(a.datetime || '').localeCompare(String(b.datetime || '')); });
  } else if (sortKey === 'title') {
    data.sort(function(a, b) { return String(a.title || '').localeCompare(String(b.title || '')); });
  } else {
    // datetime_desc (default)
    data.sort(function(a, b) { return String(b.datetime || '').localeCompare(String(a.datetime || '')); });
  }

  return { books: data.slice(0, limit || 500) };
}

// ==============================
// AIチャット
// ==============================
function chat(message) {
  var settings = getSettings();
  var apiKey = settings['gemini_api_key'];
  if (!apiKey) return { error: 'Gemini APIキーが設定されていません。settingsシートを確認してください。' };

  var today = jstToday_();
  var characterPrompt = settings['character_prompt'] || SETTINGS_DEFAULTS[2][1];

  // コンテキスト構築
  var diaryData = getSheetData_('diary');
  diaryData.sort(function(a, b) { return String(b.datetime || '').localeCompare(String(a.datetime || '')); });
  var recentEntries = diaryData.slice(0, 10);

  var todoData = getSheetData_('todo');
  var openTodos = todoData.filter(function(t) { return t.status === 'open'; }).slice(0, 10);

  var booksData = getSheetData_('books');
  booksData.sort(function(a, b) { return String(b.datetime || '').localeCompare(String(a.datetime || '')); });
  var recentBooks = booksData.slice(0, 5);

  // カレンダー予定
  var calendarText = '';
  try {
    var cal = CalendarApp.getDefaultCalendar();
    var todayDate = new Date(today + 'T00:00:00+09:00');
    var dayAfter = new Date(todayDate.getTime() + 2 * 86400000);
    var events = cal.getEvents(todayDate, dayAfter);
    if (events.length > 0) {
      var evTexts = events.map(function(ev) {
        var time = ev.isAllDayEvent() ? '終日' : Utilities.formatDate(ev.getStartTime(), 'Asia/Tokyo', 'HH:mm');
        var date = Utilities.formatDate(ev.getStartTime(), 'Asia/Tokyo', 'yyyy-MM-dd');
        return '- ' + date + ' ' + time + ' ' + ev.getTitle();
      });
      calendarText = '\n\n今日〜明日のカレンダー予定:\n' + evTexts.join('\n');
    }
  } catch (e) {}

  var contextText = '今日の日付: ' + today + '\n\n' +
    '最近の日記（最新10件）:\n' +
    (recentEntries.length ? recentEntries.map(function(e) {
      return '- ' + e.datetime + ' [mood:' + (e.mood || '-') + '] [tag:' + (e.tag || 'なし') + '] ' + (e.text || '');
    }).join('\n') : 'なし') +
    '\n\n未完了ToDo:\n' +
    (openTodos.length ? openTodos.map(function(t) {
      return '- [' + (t.priority || 'mid') + '] ' + (t.text || '') + ' (期限:' + (t.due || 'なし') + ') [tag:' + (t.tag || 'なし') + ']';
    }).join('\n') : 'なし') +
    '\n\n最近の読書:\n' +
    (recentBooks.length ? recentBooks.map(function(b) {
      return '- ' + (b.title || '') + '（' + (b.author || '') + '）★' + (b.rating || '-') + ' [' + (b.status || '') + '] ' + (b.note || '');
    }).join('\n') : 'なし') +
    calendarText;

  try {
    var payload = {
      contents: [
        { role: 'user', parts: [{ text: characterPrompt + '\n\n【コンテキスト】\n' + contextText + '\n\n【ユーザーのメッセージ】\n' + message }] }
      ],
      generationConfig: { maxOutputTokens: 2048, temperature: 0.7 }
    };

    var res = UrlFetchApp.fetch(
      'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=' + apiKey,
      {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      }
    );

    var json = JSON.parse(res.getContentText());

    if (res.getResponseCode() !== 200) {
      var errMsg = (json.error && json.error.message) || '不明なエラー';
      return { error: 'Gemini APIエラー: ' + errMsg };
    }

    var reply = 'うまく答えられなかったよ…ごめんね！';
    if (json.candidates && json.candidates[0] && json.candidates[0].content && json.candidates[0].content.parts && json.candidates[0].content.parts[0]) {
      reply = json.candidates[0].content.parts[0].text;
    }

    return { reply: reply, charName: settings['character_name'] || 'ピアちゃん' };
  } catch (e) {
    return { error: 'Gemini API通信エラー: ' + e.message };
  }
}

// ==============================
// 週次レポート
// ==============================
function sendWeeklyReport() {
  var settings = getSettings();
  var email = settings['report_email'];
  if (!email) return;

  var apiKey = settings['gemini_api_key'];
  var today = jstToday_();

  // 今週の月曜〜日曜を計算
  var jst = new Date(Date.now() + 9 * 60 * 60 * 1000);
  var day = jst.getUTCDay();
  var toMon = day === 0 ? -6 : 1 - day;
  var mon = new Date(jst);
  mon.setUTCDate(jst.getUTCDate() + toMon);
  var from = mon.toISOString().slice(0, 10);
  var nextMon = new Date(from + 'T00:00:00Z');
  nextMon.setUTCDate(nextMon.getUTCDate() + 7);
  var to = nextMon.toISOString().slice(0, 10);

  var lastDay = new Date(new Date(to + 'T00:00:00Z').getTime() - 86400000).toISOString().slice(0, 10);
  var periodLabel = from + ' 〜 ' + lastDay;

  // データ収集
  var diaryData = getSheetData_('diary');
  var todoData = getSheetData_('todo');
  var booksData = getSheetData_('books');

  var weekEntries = diaryData.filter(function(e) { return e.datetime >= from && e.datetime < to; });
  var weekDoneTodos = todoData.filter(function(t) { return t.status === 'done' && t.done_at >= from && t.done_at < to; });
  var weekOpenTodos = todoData.filter(function(t) { return t.status !== 'done' && t.due && t.due < to; });
  var weekBooks = booksData.filter(function(b) { return b.status === 'done' && b.datetime >= from && b.datetime < to; });

  // 集計
  var moodEntries = weekEntries.filter(function(e) { return e.mood && e.mood > 0; });
  var avgMood = moodEntries.length ? Math.round((moodEntries.reduce(function(s, e) { return s + Number(e.mood); }, 0) / moodEntries.length) * 10) / 10 : null;

  var stats = {
    entryCount: weekEntries.length,
    avgMood: avgMood,
    todoCompleted: weekDoneTodos.length,
    todoRemaining: weekOpenTodos.length,
    booksFinished: weekBooks.length
  };

  // 先週データ（比較用）
  var lastFrom = new Date(new Date(from + 'T00:00:00Z').getTime() - 7 * 86400000).toISOString().slice(0, 10);
  var lastTo = from;
  var lastEntries = diaryData.filter(function(e) { return e.datetime >= lastFrom && e.datetime < lastTo; });
  var lastMoodEntries = lastEntries.filter(function(e) { return e.mood && e.mood > 0; });
  var lastWeek = {
    entryCount: lastEntries.length,
    avgMood: lastMoodEntries.length ? Math.round((lastMoodEntries.reduce(function(s, e) { return s + Number(e.mood); }, 0) / lastMoodEntries.length) * 10) / 10 : null,
    todoDone: todoData.filter(function(t) { return t.status === 'done' && t.done_at >= lastFrom && t.done_at < lastTo; }).length,
    bookCount: booksData.filter(function(b) { return b.status === 'done' && b.datetime >= lastFrom && b.datetime < lastTo; }).length
  };

  // 来週のカレンダー予定
  var nextWeekEvents = [];
  try {
    var cal = CalendarApp.getDefaultCalendar();
    var nextFrom = new Date(to + 'T00:00:00+09:00');
    var nextTo = new Date(nextFrom.getTime() + 7 * 86400000);
    var events = cal.getEvents(nextFrom, nextTo);
    events.forEach(function(ev) {
      nextWeekEvents.push({
        date: Utilities.formatDate(ev.getStartTime(), 'Asia/Tokyo', 'yyyy-MM-dd'),
        startTime: ev.isAllDayEvent() ? '終日' : Utilities.formatDate(ev.getStartTime(), 'Asia/Tokyo', 'HH:mm'),
        endTime: ev.isAllDayEvent() ? '' : Utilities.formatDate(ev.getEndTime(), 'Asia/Tokyo', 'HH:mm'),
        title: ev.getTitle(),
        location: ev.getLocation() || ''
      });
    });
  } catch (e) {}

  // 来週期限のToDo
  var nextWeekTodos = todoData.filter(function(t) {
    return t.status !== 'done' && t.due >= to && t.due < new Date(new Date(to + 'T00:00:00Z').getTime() + 7 * 86400000).toISOString().slice(0, 10);
  });

  // Geminiコメント生成
  var piaComment = { summary: '', todo_comment: '', book_comment: '', next_week_advice: [] };
  if (apiKey) {
    try {
      piaComment = generateReportComment_(apiKey, {
        entries: weekEntries,
        doneTodos: weekDoneTodos,
        openTodos: weekOpenTodos,
        books: weekBooks,
        lastWeek: lastWeek,
        nextWeekEvents: nextWeekEvents,
        nextWeekTodos: nextWeekTodos
      });
    } catch (e) {
      piaComment.summary = 'コメントの生成に失敗しました。';
    }
  }

  // HTMLメール生成
  var emailHtml = buildReportHtml_(stats, piaComment, lastWeek, weekBooks, periodLabel);
  var subject = '🐾 ピアちゃんの週次レポート（' + periodLabel + '）';

  MailApp.sendEmail({
    to: email,
    subject: subject,
    htmlBody: emailHtml
  });
}

function generateReportComment_(apiKey, data) {
  var prompt = 'あなたは「ピアちゃん」。もちもちしたピンクのゆるキャラ。おっとりした口調で話す。褒めるときはしっかり褒める。ちょっとおせっかいだけど押しつけがましくない。絵文字を適度に使う（1〜2個/セクション）。\n\n' +
    '以下のデータをもとに週次レポートのコメントをJSON形式で生成してください。\n\n' +
    '## 出力フォーマット\n' +
    '{"summary":"今週の一言サマリー。2文以内","todo_comment":"ToDo消化率へのコメント。2〜3文","book_comment":"読書に関するコメント。1〜2文","next_week_advice":["提案1","提案2","提案3"]}\n\n' +
    '## ルール\n- 各コメントは短く。1セクション3文以内。\n- 数字の羅列はしない。\n- ネガティブなことも受け止めるが、説教しない。\n- JSONのみ出力。\n\n' +
    '## 今週のデータ\n' +
    '### 日記（' + data.entries.length + '件）\n' +
    (data.entries.length ? data.entries.map(function(e) { return e.datetime + ' mood:' + (e.mood || '-') + ' "' + (e.text || '').slice(0, 100) + '"'; }).join('\n') : 'なし') + '\n\n' +
    '### ToDo完了（' + data.doneTodos.length + '件）\n' +
    (data.doneTodos.length ? data.doneTodos.map(function(t) { return '✓ ' + t.text; }).join('\n') : 'なし') + '\n\n' +
    '### ToDo未完了（' + data.openTodos.length + '件）\n' +
    (data.openTodos.length ? data.openTodos.map(function(t) { return '□ ' + t.text + ' 期限:' + (t.due || 'なし'); }).join('\n') : 'なし') + '\n\n' +
    '### 読書（' + data.books.length + '冊）\n' +
    (data.books.length ? data.books.map(function(b) { return '「' + b.title + '」★' + (b.rating || '-') + ' ' + (b.note || ''); }).join('\n') : 'なし') + '\n\n' +
    '### 先週の数値（比較用）\n' +
    '日記:' + data.lastWeek.entryCount + '件 / mood:' + (data.lastWeek.avgMood || '-') + ' / ToDo完了:' + data.lastWeek.todoDone + '件 / 読書:' + data.lastWeek.bookCount + '冊\n\n' +
    '### 来週のカレンダー予定\n' +
    (data.nextWeekEvents.length ? data.nextWeekEvents.map(function(e) { return e.date + ' ' + e.startTime + ' ' + e.title; }).join('\n') : 'なし') + '\n\n' +
    '### 来週期限のToDo\n' +
    (data.nextWeekTodos.length ? data.nextWeekTodos.map(function(t) { return '□ ' + t.text + ' 期限:' + t.due; }).join('\n') : 'なし');

  var res = UrlFetchApp.fetch(
    'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=' + apiKey,
    {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }] }),
      muteHttpExceptions: true
    }
  );
  var json = JSON.parse(res.getContentText());
  var raw = (json.candidates && json.candidates[0] && json.candidates[0].content && json.candidates[0].content.parts && json.candidates[0].content.parts[0]) ? json.candidates[0].content.parts[0].text : '{}';
  var cleaned = raw.replace(/^```(?:json)?\s*/i, '').replace(/\s*```\s*$/, '').trim();
  try {
    return JSON.parse(cleaned);
  } catch (e) {
    return { summary: raw.slice(0, 200), todo_comment: '', book_comment: '', next_week_advice: [] };
  }
}

function buildReportHtml_(stats, piaComment, lastWeek, books, periodLabel) {
  var moodEmoji = function(m) { return ['', '😢', '😞', '😐', '🙂', '😊', '🤩'][Math.round(m)] || '➖'; };

  var delta = function(cur, prev) {
    if (prev == null || cur == null) return '';
    var d = Math.round((cur - prev) * 10) / 10;
    if (d === 0) return '<span style="color:#9ca3af;font-size:11px;">±0</span>';
    var good = d > 0;
    return '<span style="color:' + (good ? '#10b981' : '#ef4444') + ';font-size:11px;">' + (d > 0 ? '↑' : '↓') + Math.abs(d) + '</span>';
  };

  var card = function(emoji, label, value, deltaHtml) {
    return '<td style="background:#FFF0F5;border-radius:10px;padding:12px;text-align:center;width:33%;">' +
      '<div style="font-size:20px;">' + emoji + '</div>' +
      '<div style="font-size:11px;color:#7A9490;">' + label + '</div>' +
      '<div style="font-size:18px;font-weight:700;color:#2D3B36;">' + value + '</div>' +
      (deltaHtml || '') + '</td>';
  };

  var adviceItems = (piaComment.next_week_advice || []).map(function(a) {
    return '<li style="margin-bottom:6px;font-size:13px;color:#2D3B36;">' + a + '</li>';
  }).join('');

  return '<!DOCTYPE html><html><head><meta charset="utf-8"></head>' +
    '<body style="margin:0;padding:0;background:#fafafa;font-family:\'Helvetica Neue\',Arial,sans-serif;">' +
    '<div style="max-width:560px;margin:24px auto;background:#fff;border-radius:16px;overflow:hidden;box-shadow:0 4px 24px rgba(0,0,0,0.08);">' +
    '<div style="background:linear-gradient(135deg,#E8A0BF,#957DAD);padding:28px 24px;text-align:center;">' +
    '<div style="font-size:32px;">🐾</div>' +
    '<h1 style="margin:8px 0 0;color:#fff;font-size:20px;font-weight:700;">ピアちゃんの週次レポート</h1>' +
    '<p style="margin:6px 0 0;color:rgba(255,255,255,0.85);font-size:13px;">' + periodLabel + '</p>' +
    '</div>' +
    '<div style="padding:20px 24px;">' +
    '<div style="background:#FFF0F5;border-radius:12px;padding:16px;margin-bottom:20px;border-left:4px solid #E8A0BF;">' +
    '<div style="font-weight:600;color:#957DAD;margin-bottom:6px;">🐾 今週のサマリー</div>' +
    '<p style="margin:0;color:#2D3B36;font-size:14px;line-height:1.7;">' + (piaComment.summary || '') + '</p>' +
    '</div>' +
    '<table style="width:100%;border-collapse:separate;border-spacing:8px;margin-bottom:8px;"><tr>' +
    card('📝', '日記', stats.entryCount + '件', delta(stats.entryCount, lastWeek.entryCount)) +
    card(stats.avgMood ? moodEmoji(stats.avgMood) : '➖', '平均mood', stats.avgMood || '—', delta(stats.avgMood, lastWeek.avgMood)) +
    card('✅', 'ToDo完了', stats.todoCompleted + '件', delta(stats.todoCompleted, lastWeek.todoDone)) +
    '</tr><tr>' +
    card('📋', 'ToDo残', stats.todoRemaining + '件', '') +
    card('📚', '読書', stats.booksFinished + '冊', delta(stats.booksFinished, lastWeek.bookCount)) +
    '<td style="background:#FFF0F5;border-radius:10px;padding:12px;text-align:center;width:33%;"></td>' +
    '</tr></table>' +
    (piaComment.todo_comment ? '<div style="background:#f9fafb;border-radius:10px;padding:14px;margin-top:16px;"><div style="font-weight:600;color:#957DAD;margin-bottom:4px;">✅ ToDo</div><p style="margin:0;font-size:13px;color:#2D3B36;line-height:1.7;">' + piaComment.todo_comment + '</p></div>' : '') +
    (books.length ? '<div style="background:#f9fafb;border-radius:10px;padding:14px;margin-top:12px;"><div style="font-weight:600;color:#957DAD;margin-bottom:4px;">📚 読書</div>' + books.map(function(b) { return '<div style="font-size:13px;color:#2D3B36;">「' + b.title + '」' + '★'.repeat(b.rating || 0) + '</div>'; }).join('') + (piaComment.book_comment ? '<p style="margin:6px 0 0;font-size:13px;color:#7A9490;line-height:1.7;">' + piaComment.book_comment + '</p>' : '') + '</div>' : '') +
    (adviceItems ? '<div style="background:#FFF0F5;border-radius:10px;padding:14px;margin-top:16px;"><div style="font-weight:600;color:#957DAD;margin-bottom:8px;">🌸 来週へのアドバイス</div><ul style="margin:0;padding-left:18px;">' + adviceItems + '</ul></div>' : '') +
    '</div>' +
    '<div style="padding:12px 24px;text-align:center;color:#9ca3af;font-size:11px;border-top:1px solid #f3f4f6;">Life OS — あなたの毎日をサポート 🐾</div>' +
    '</div></body></html>';
}

// ==============================
// トリガー設定
// ==============================
function setupWeeklyTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(t) {
    if (t.getHandlerFunction() === 'sendWeeklyReport') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sendWeeklyReport')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.SUNDAY)
    .atHour(9)
    .create();
}
