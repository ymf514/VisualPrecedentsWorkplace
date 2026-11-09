/**
 * Humor Precedents: Google Sheet storage layer.
 *
 * 1. In your Google Sheet: Extensions > Apps Script, paste this whole file, save.
 * 2. Choose the function `setup` in the toolbar and click Run. Approve access.
 *    The execution log prints your edit token. Copy it.
 * 3. Deploy > New deployment > Web app.
 *    Execute as: Me. Who has access: Anyone. Copy the /exec URL.
 *
 * Reading is public so visitors of the website can see the archive.
 * Every write needs the token, which lives in Script Properties and never in this file.
 */

var KEY_COLUMNS = { Images: 'id', Tags: 'key', Settings: 'key' };

var BASE_HEADERS = {
  Images: [
    'id', 'file', 'original_file', 'width', 'height', 'status',
    'title_zh', 'title_en',
    'alias_zh', 'alias_en',
    'see_zh', 'see_en', 'grab_zh', 'grab_en', 'reply_zh', 'reply_en',
    'why_zh', 'why_en', 'offers_zh', 'offers_en',
    'creator_zh', 'creator_en', 'year_zh', 'year_en', 'place', 'source_type', 'source_url',
    'context_zh', 'context_en',
    'tag_medium_zh', 'tag_medium_en',
    'tag_language_zh', 'tag_language_en',
    'tag_punchline_zh', 'tag_punchline_en',
    'en_locked', 'updated'
  ],
  Tags: ['key', 'dim', 'dim_zh', 'dim_en', 'zh', 'en', 'order'],
  Settings: ['key', 'value']
};

var SEED_IMAGES = [
  ['VP-001', '227a9972ca58943a5759219325d53e89.jpg', 1365, 1280],
  ['VP-002', '52c205e3b0b67dc57d02d1a567f14ba1.jpg', 1260, 707],
  ['VP-003', '7e6a28d79aa8ec89520dc4fe9f8fae1c.jpg', 1280, 1706],
  ['VP-004', 'c6e9084dd92fb35da6519c3ec362c8b4.jpg', 1284, 660],
  ['VP-005', 'cf6170ff90c0d9abb5b6d096367a57c6.jpg', 1260, 664],
  ['VP-006', 'd7ed06f9c2c6ceb2ca5c72cdec6cd75c.jpg', 1260, 1914],
  ['VP-007', 'ed7a4d65ca286d2b4b5c6209489d7287.jpg', 1260, 1213],
  ['VP-008', 'ef6c17ad41de693fa763616992e25fa1.jpg', 1260, 1992]
];

var SEED_TAGS = {
  medium: ['媒介', 'Medium', [
    ['图像', 'Image'], ['视频', 'Video'], ['漫画', 'Comic'], ['文字', 'Text'],
    ['表情包', 'Meme'], ['装置', 'Installation'], ['行为', 'Performance']
  ]],
  language: ['语言', 'Language', [
    ['中文', 'Chinese'], ['英文', 'English'], ['方言', 'Dialect'], ['无文字', 'Wordless']
  ]],
  punchline: ['笑点', 'Punchline', [
    ['反差', 'Contrast'], ['谐音', 'Pun'], ['荒诞', 'Absurdity'], ['自嘲', 'Self-deprecation'],
    ['夸张', 'Exaggeration'], ['错位', 'Mismatch'], ['反讽', 'Irony']
  ]]
};

var SEED_SETTINGS = {
  thesis_zh: '收集能体现我本人对幽默及幽默表达的理解的东西。',
  thesis_en: 'Collecting what embodies my own understanding of humor and how it is expressed.',
  target: '100',
  last_id: '8'
};

function setup() {
  var book = SpreadsheetApp.getActive();
  Object.keys(BASE_HEADERS).forEach(function (name) {
    var sheet = book.getSheetByName(name) || book.insertSheet(name);
    ensureHeaders_(sheet, BASE_HEADERS[name]);
    sheet.setFrozenRows(1);
  });

  var ops = [];
  if (readSheet_('Images').length === 0) {
    SEED_IMAGES.forEach(function (row) {
      ops.push({ op: 'upsert', sheet: 'Images', key: row[0], patch: {
        file: row[0] + '.jpg', original_file: row[1],
        width: row[2], height: row[3], status: 'active'
      } });
    });
  }
  if (readSheet_('Tags').length === 0) {
    var order = 0;
    Object.keys(SEED_TAGS).forEach(function (dim) {
      var seed = SEED_TAGS[dim];
      ops.push({ op: 'upsert', sheet: 'Tags', key: dim + '::', patch: {
        dim: dim, dim_zh: seed[0], dim_en: seed[1], zh: '', en: '', order: order++
      } });
      seed[2].forEach(function (tag) {
        ops.push({ op: 'upsert', sheet: 'Tags', key: dim + '::' + tag[0], patch: {
          dim: dim, zh: tag[0], en: tag[1], order: order++
        } });
      });
    });
  }
  if (readSheet_('Settings').length === 0) {
    Object.keys(SEED_SETTINGS).forEach(function (key) {
      ops.push({ op: 'upsert', sheet: 'Settings', key: key, patch: { value: SEED_SETTINGS[key] } });
    });
  }
  applyBatch_(ops);

  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty('TOKEN')) props.setProperty('TOKEN', Utilities.getUuid());
  Logger.log('Your edit token: ' + props.getProperty('TOKEN'));
}

function doGet() {
  try {
    var settings = {};
    readSheet_('Settings').forEach(function (row) { settings[row.key] = row.value; });
    return json_({
      ok: true,
      images: readSheet_('Images'),
      tags: readSheet_('Tags'),
      settings: settings
    });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var expected = PropertiesService.getScriptProperties().getProperty('TOKEN');
    if (!expected || body.token !== expected) return json_({ ok: false, error: 'bad_token' });

    if (body.action === 'verify') return json_({ ok: true });
    if (body.action === 'translate') return json_({ ok: true, texts: translate_(body) });
    if (body.action === 'batch') {
      var lock = LockService.getScriptLock();
      lock.waitLock(20000);
      try {
        applyBatch_(body.ops || []);
      } finally {
        lock.releaseLock();
      }
      return json_({ ok: true });
    }
    return json_({ ok: false, error: 'unknown_action' });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

function translate_(body) {
  return (body.texts || []).map(function (text) {
    if (!String(text || '').trim()) return '';
    return LanguageApp.translate(String(text), body.from, body.to);
  });
}

function applyBatch_(ops) {
  var tables = {};
  ops.forEach(function (item) {
    if (!Object.prototype.hasOwnProperty.call(KEY_COLUMNS, item.sheet)) throw new Error('bad_sheet');
    if (!item.key) throw new Error('bad_key');
    var table = tables[item.sheet] || (tables[item.sheet] = openTable_(item.sheet));
    if (item.op === 'upsert') upsert_(table, String(item.key), item.patch || {});
    else if (item.op === 'delete' && item.sheet === 'Tags') remove_(table, String(item.key));
    else throw new Error('bad_op');
  });
}

function openTable_(name) {
  var sheet = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sheet) throw new Error('missing_sheet_' + name + '_run_setup');
  var headers = ensureHeaders_(sheet, BASE_HEADERS[name]);
  var last = sheet.getLastRow();
  var keys = last > 1 ? sheet.getRange(2, 1, last - 1, 1).getDisplayValues() : [];
  return { sheet: sheet, headers: headers, keys: keys.map(function (r) { return String(r[0]); }) };
}

function upsert_(table, key, patch) {
  var fresh = Object.keys(patch).filter(function (field) {
    if (!/^[a-z0-9_]+$/.test(field)) throw new Error('bad_field');
    return table.headers.indexOf(field) === -1;
  });
  if (fresh.length) table.headers = ensureHeaders_(table.sheet, table.headers.concat(fresh));

  var index = table.keys.indexOf(key);
  var isNew = index === -1;
  if (isNew) {
    table.keys.push(key);
    index = table.keys.length - 1;
  }
  var rowNumber = index + 2;
  var sheet = table.sheet;
  if (rowNumber > sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(), 50);

  var range = sheet.getRange(rowNumber, 1, 1, table.headers.length);
  var values = isNew
    ? table.headers.map(function () { return ''; })
    : range.getDisplayValues()[0];
  values[0] = key;
  Object.keys(patch).forEach(function (field) {
    var col = table.headers.indexOf(field);
    if (col > 0) values[col] = String(patch[field] == null ? '' : patch[field]);
  });
  range.setNumberFormat('@').setValues([values]);
}

function remove_(table, key) {
  var index = table.keys.indexOf(key);
  if (index === -1) return;
  table.sheet.deleteRow(index + 2);
  table.keys.splice(index, 1);
}

function ensureHeaders_(sheet, wanted) {
  var maxCols = sheet.getMaxColumns();
  var used = sheet.getLastColumn();
  var existing = used ? sheet.getRange(1, 1, 1, used).getDisplayValues()[0].filter(String) : [];
  var headers = existing.slice();
  wanted.forEach(function (field) {
    if (headers.indexOf(field) === -1) headers.push(field);
  });
  if (headers.length > existing.length) {
    if (headers.length > maxCols) sheet.insertColumnsAfter(maxCols, headers.length - maxCols);
    sheet.getRange(1, 1, 1, headers.length).setNumberFormat('@').setValues([headers]);
  }
  return headers;
}

function readSheet_(name) {
  var sheet = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sheet || sheet.getLastRow() < 2) return [];
  var data = sheet.getDataRange().getDisplayValues();
  var headers = data.shift();
  return data.filter(function (row) { return row[0] !== ''; }).map(function (row) {
    var item = {};
    headers.forEach(function (field, col) { if (field) item[field] = row[col]; });
    return item;
  });
}

function json_(value) {
  return ContentService.createTextOutput(JSON.stringify(value))
    .setMimeType(ContentService.MimeType.JSON);
}
