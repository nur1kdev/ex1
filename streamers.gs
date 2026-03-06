// ============================================================
//  АВТООБНОВЛЕНИЕ AVG ONLINE + FOLLOWERS  v5
//  TwitchTracker API → AVG viewers за 30 дней (Twitch)
//  Twitch API        → Followers (Twitch)
//  Kick API          → Followers + онлайн (Kick)
// ============================================================

// ── ТВОИ КЛЮЧИ ──────────────────────────────────────────────
const TWITCH_CLIENT_ID     = ''; // fallback: можно оставить пустым и хранить в Script Properties
const TWITCH_CLIENT_SECRET = ''; // fallback: можно оставить пустым и хранить в Script Properties
const SCRIPT_PROP_TWITCH_CLIENT_ID = 'TWITCH_CLIENT_ID';
const SCRIPT_PROP_TWITCH_CLIENT_SECRET = 'TWITCH_CLIENT_SECRET';
// ───────────────────────────────────────────────────────────

// ── НАСТРОЙКИ ТАБЛИЦЫ ───────────────────────────────────────
const SHEET_NAME     = 'main list';
const COL_URL        = 2;  // B — ссылка на стримера
const COL_AVG_ONLINE = 5;  // E — AVG Online
const COL_FOLLOWERS  = 6;  // F — Followers
const START_ROW      = 2;
// ───────────────────────────────────────────────────────────



function getTwitchCredentials() {
  const props = PropertiesService.getScriptProperties();
  const propClientId = props.getProperty(SCRIPT_PROP_TWITCH_CLIENT_ID);
  const propClientSecret = props.getProperty(SCRIPT_PROP_TWITCH_CLIENT_SECRET);

  const clientId = propClientId || TWITCH_CLIENT_ID;
  const clientSecret = propClientSecret || TWITCH_CLIENT_SECRET;

  if (!clientId || !clientSecret) return null;
  return { clientId, clientSecret };
}

function setTwitchCredentials(clientId, clientSecret) {
  if (!clientId || !clientSecret) {
    throw new Error('Передай TWITCH_CLIENT_ID и TWITCH_CLIENT_SECRET');
  }

  const props = PropertiesService.getScriptProperties();
  props.setProperty(SCRIPT_PROP_TWITCH_CLIENT_ID, clientId);
  props.setProperty(SCRIPT_PROP_TWITCH_CLIENT_SECRET, clientSecret);
}

// ── ГЛАВНАЯ ФУНКЦИЯ ─────────────────────────────────────────

function updateAllStreamers() {
  const sheet   = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow();
  const token   = getTwitchToken();

  if (!token) {
    SpreadsheetApp.getUi().alert('❌ Не удалось получить Twitch токен. Проверь Client ID и Secret.');
    return;
  }

  let updated = 0;
  let errors  = [];

  for (let row = START_ROW; row <= lastRow; row++) {
    const url = sheet.getRange(row, COL_URL).getValue();
    if (!url || typeof url !== 'string') continue;

    const urlLower = url.toLowerCase();

    try {
      if (urlLower.includes('twitch.tv')) {
        const username = extractUsername(url, 'twitch.tv');
        if (!username) continue;

        const avgOnline = fetchTwitchTrackerAvg(username);
        const followers = fetchTwitchFollowers(username, token);

        if (avgOnline !== null) sheet.getRange(row, COL_AVG_ONLINE).setValue(avgOnline);
        if (followers !== null) sheet.getRange(row, COL_FOLLOWERS).setValue(followers);

        updated++;
        Utilities.sleep(1500); // пауза между запросами к TwitchTracker

      } else if (urlLower.includes('kick.com')) {
        const username = extractUsername(url, 'kick.com');
        if (!username) continue;

        const data = fetchKick(username);
        if (data) {
          if (data.avgOnline !== null) sheet.getRange(row, COL_AVG_ONLINE).setValue(data.avgOnline);
          if (data.followers !== null) sheet.getRange(row, COL_FOLLOWERS).setValue(data.followers);
          updated++;
        }
        Utilities.sleep(800);
      }

    } catch(e) {
      errors.push(`Строка ${row}: ${e.message}`);
    }
  }

  let msg = `✅ Обновлено: ${updated} стримеров`;
  if (errors.length > 0) {
    msg += `\n\n⚠️ Ошибки (${errors.length}):\n` + errors.slice(0, 10).join('\n');
  }
  SpreadsheetApp.getUi().alert(msg);
}


// ── ОБНОВИТЬ ОДНУ СТРОКУ ─────────────────────────────────────

function updateSelectedRow() {
  // Запрашиваем номер строки вручную через диалог
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Берём текущую выделенную строку
  const selection = sheet.getSelection();
  const row = selection.getActiveRange().getRow();

  // Показываем какой стример будет обновлён
  const urlCheck = sheet.getRange(row, COL_URL).getValue();
  if (!urlCheck) { ui.alert('В строке ' + row + ' нет ссылки. Кликни на ячейку со ссылкой стримера.'); return; }

  const confirm = ui.alert('Обновить стримера?', urlCheck, ui.ButtonSet.OK_CANCEL);
  if (confirm !== ui.Button.OK) return;


  if (row < START_ROW) {
    SpreadsheetApp.getUi().alert('Выбери строку с данными стримера');
    return;
  }

  const url = sheet.getRange(row, COL_URL).getValue();
  if (!url) { SpreadsheetApp.getUi().alert('В строке нет ссылки'); return; }

  const urlLower = url.toLowerCase();
  let avgOnline = null;
  let followers = null;

  if (urlLower.includes('twitch.tv')) {
    const username = extractUsername(url, 'twitch.tv');
    const token    = getTwitchToken();
    avgOnline = fetchTwitchTrackerAvg(username);
    followers = fetchTwitchFollowers(username, token);

  } else if (urlLower.includes('kick.com')) {
    const username = extractUsername(url, 'kick.com');
    const data     = fetchKick(username);
    if (data) {
      avgOnline = data.avgOnline;
      followers = data.followers;
    }
  }

  if (avgOnline !== null) sheet.getRange(row, COL_AVG_ONLINE).setValue(avgOnline);
  if (followers !== null) sheet.getRange(row, COL_FOLLOWERS).setValue(followers);

  SpreadsheetApp.getUi().alert(
    `✅ Обновлено!\nAVG Online (30 дней): ${avgOnline ?? 'не найдено'}\nFollowers: ${followers ?? 'не найдено'}`
  );
}


// ── TWITCHTRACKER API — AVG VIEWERS ─────────────────────────
// Официальный endpoint: /api/channels/summary/[channel_name]

function fetchTwitchTrackerAvg(username) {
  const apiUrl = `https://twitchtracker.com/api/channels/summary/${username.toLowerCase()}`;

  // Шаг 1: посещаем страницу стримера чтобы получить cookies
  let cookieHeader = '';
  try {
    const pageResp = UrlFetchApp.fetch(`https://twitchtracker.com/${username.toLowerCase()}`, {
      method: 'get',
      headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36' },
      muteHttpExceptions: true
    });
    const setCookie = pageResp.getAllHeaders()['Set-Cookie'];
    if (setCookie) {
      const cookieArr = Array.isArray(setCookie) ? setCookie : [setCookie];
      cookieHeader = cookieArr.map(c => c.split(';')[0]).join('; ');
    }
  } catch(e) {}

  // Шаг 2: запрашиваем API с cookies
  const options = {
    method: 'get',
    headers: {
      'User-Agent':       'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
      'Accept':           'application/json, text/plain, */*',
      'Accept-Language':  'en-US,en;q=0.9',
      'Referer':          `https://twitchtracker.com/${username.toLowerCase()}`,
      'Origin':           'https://twitchtracker.com',
      'X-Requested-With': 'XMLHttpRequest',
      'Cookie':           cookieHeader,
    },
    muteHttpExceptions: true
  };

  const resp = UrlFetchApp.fetch(apiUrl, options);
  const code = resp.getResponseCode();
  const body = resp.getContentText();

  Logger.log(`TwitchTracker [${username}] HTTP ${code}: ${body.substring(0, 300)}`);

  if (code !== 200) return null;

  try {
    const data = JSON.parse(body);
    return data.avg_viewers ?? data.avg_cviewers ?? data.average_viewers ?? null;
  } catch(e) {
    Logger.log(`TwitchTracker parse error [${username}]: ${e.message}`);
    return null;
  }
}


// ── TWITCH API — FOLLOWERS ───────────────────────────────────

function getTwitchToken() {
  const cache = PropertiesService.getScriptProperties();
  const saved = cache.getProperty('twitch_token');
  const exp   = cache.getProperty('twitch_token_exp');

  if (saved && exp && Date.now() < parseInt(exp, 10)) return saved;

  const creds = getTwitchCredentials();
  if (!creds) return null;

  const resp = UrlFetchApp.fetch('https://id.twitch.tv/oauth2/token', {
    method: 'post',
    payload: {
      client_id:     creds.clientId,
      client_secret: creds.clientSecret,
      grant_type:    'client_credentials'
    },
    muteHttpExceptions: true
  });

  const data = JSON.parse(resp.getContentText());
  if (!data.access_token) return null;

  cache.setProperty('twitch_token', data.access_token);
  cache.setProperty('twitch_token_exp', (Date.now() + 30 * 24 * 3600 * 1000).toString());

  return data.access_token;
}

function fetchTwitchFollowers(username, token) {
  const creds = getTwitchCredentials();
  if (!creds || !token) return null;

  const userResp = UrlFetchApp.fetch(
    `https://api.twitch.tv/helix/users?login=${encodeURIComponent(username)}`,
    {
      headers: {
        'Client-ID':     creds.clientId,
        'Authorization': `Bearer ${token}`
      },
      muteHttpExceptions: true
    }
  );

  const userData = JSON.parse(userResp.getContentText());
  if (!userData.data || userData.data.length === 0) return null;

  const userId = userData.data[0].id;

  const fResp = UrlFetchApp.fetch(
    `https://api.twitch.tv/helix/channels/followers?broadcaster_id=${userId}&first=1`,
    {
      headers: {
        'Client-ID':     creds.clientId,
        'Authorization': `Bearer ${token}`
      },
      muteHttpExceptions: true
    }
  );

  const fData = JSON.parse(fResp.getContentText());
  return fData.total ?? null;
}


// ── KICK — AVG VIEWERS (StreamCharts) + FOLLOWERS ───────────

function fetchKick(username) {
  const avgOnline = fetchStreamChartsKickAvg(username);
  const followers = fetchKickFollowers(username);
  return { avgOnline, followers };
}

function fetchStreamChartsKickAvg(username) {
  // AeroKick — данные встроены в HTML как SSR, парсим avg_viewers из страницы
  const pageUrl = `https://aerokick.app/stats/channels/${username.toLowerCase()}?range=month`;

  const resp = UrlFetchApp.fetch(pageUrl, {
    method: 'get',
    headers: {
      'User-Agent':      'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
      'Accept':          'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
      'Accept-Language': 'en-US,en;q=0.9',
    },
    muteHttpExceptions: true
  });

  const code = resp.getResponseCode();
  const body = resp.getContentText();

  Logger.log(`AeroKick [${username}] HTTP ${code}: ${body.substring(0, 200)}`);

  if (code !== 200) return null;

  try {
    // Данные встроены в HTML как SSR в блоках streamController.enqueue()
    // Структура в конце второго блока: "TopCategories",[...], "Counter-Strike 2", "value", 108.91562]
    // Ищем "\"value\"," за которым сразу идёт число — это avg по топ категории за период

    // Метод 1: "value", число — в конце данных страницы
    const valueMatch = body.match(/"value",(\d+\.?\d*)\]/);
    if (valueMatch) {
      Logger.log(`AeroKick [${username}] value match: ${valueMatch[1]}`);
      return Math.round(parseFloat(valueMatch[1]));
    }

    // Метод 2: найти среднее прямо в HTML из блока с текстом Average Viewers
    // <div class="...tabular-nums...">100.6</div>
    const tabularMatch = body.match(/tabular-nums[^>]*>(\d+\.?\d*)<\/div>/g);
    if (tabularMatch && tabularMatch.length > 0) {
      const firstNum = tabularMatch[0].match(/>(\d+\.?\d*)</);
      if (firstNum) {
        Logger.log(`AeroKick [${username}] tabular match: ${firstNum[1]}`);
        return Math.round(parseFloat(firstNum[1]));
      }
    }

    // Метод 3: найти "Average Viewers" и взять число рядом
    const avgSection = body.match(/Average Viewers[\s\S]{0,300}/);
    if (avgSection) {
      const num = avgSection[0].match(/(\d+\.?\d{0,2})/g);
      if (num) {
        Logger.log(`AeroKick [${username}] avg section: ${num[0]}`);
        return Math.round(parseFloat(num[0]));
      }
    }

    Logger.log(`AeroKick [${username}] no match found`);
    return null;
  } catch(e) {
    Logger.log(`AeroKick parse error [${username}]: ${e}`);
    return null;
  }
}

function extractKickFollowers(payload) {
  if (!payload || typeof payload !== 'object') return null;

  const candidates = [
    payload?.followers_count,
    payload?.data?.followers_count,
    payload?.channel?.followers_count,
    payload?.data?.channel?.followers_count,
    payload?.livestream?.channel?.followers_count,
  ];

  for (const candidate of candidates) {
    const n = Number(candidate);
    if (Number.isFinite(n)) return Math.round(n);
  }

  return null;
}

function fetchKickFollowers(username) {
  const normalized = username.toLowerCase();
  const endpoints = [
    `https://kick.com/api/v2/channels/${normalized}`,
    `https://api.kick.com/private/v1/channels/${normalized}`,
    `https://api.kick.com/public/v1/channels/${normalized}`,
  ];

  for (const endpoint of endpoints) {
    for (let attempt = 1; attempt <= 2; attempt++) {
      try {
        const resp = UrlFetchApp.fetch(endpoint, {
          headers: {
            'Accept': 'application/json',
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
            'Origin': 'https://kick.com',
            'Referer': `https://kick.com/${normalized}`,
          },
          muteHttpExceptions: true,
          followRedirects: true,
        });

        const code = resp.getResponseCode();
        if (code === 429 || code >= 500) {
          if (attempt < 2) {
            Utilities.sleep(700);
            continue;
          }
          Logger.log(`Kick followers retry failed ${code} ${endpoint}`);
          break;
        }

        if (code !== 200) {
          Logger.log(`Kick followers HTTP ${code} ${endpoint}`);
          break;
        }

        const payload = JSON.parse(resp.getContentText());
        const followers = extractKickFollowers(payload);
        if (followers !== null) {
          Logger.log(`Kick followers [${normalized}] source=${endpoint} value=${followers}`);
          return followers;
        }

        break;
      } catch (e) {
        if (attempt < 2) {
          Utilities.sleep(700);
          continue;
        }
        Logger.log(`Kick followers error ${endpoint}: ${e.message}`);
      }
    }
  }

  // HTML fallback
  try {
    const pageResp = UrlFetchApp.fetch(`https://kick.com/${normalized}`, {
      headers: { 'User-Agent': 'Mozilla/5.0' },
      muteHttpExceptions: true,
      followRedirects: true,
    });

    if (pageResp.getResponseCode() === 200) {
      const html = pageResp.getContentText();
      const m = html.match(/"followers_count"\s*:\s*(\d+)/i) || html.match(/followers_count\D+(\d+)/i);
      if (m) {
        const n = Number(m[1]);
        if (Number.isFinite(n)) {
          Logger.log(`Kick followers [${normalized}] source=html value=${n}`);
          return n;
        }
      }
    }
  } catch (e) {
    Logger.log(`Kick followers html fallback error [${normalized}]: ${e.message}`);
  }

  return null;
}


// ── УТИЛИТЫ ──────────────────────────────────────────────────

function extractUsername(url, domain) {
  const regex = new RegExp(domain.replace('.', '\\.') + '\\/([a-zA-Z0-9_]+)');
  const match = url.match(regex);
  return match ? match[1] : null;
}


// ── МЕНЮ ─────────────────────────────────────────────────────

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🎮 Стримеры')
    .addItem('Обновить всех', 'updateAllStreamers')
    .addItem('Обновить выбранную строку', 'updateSelectedRow')
    .addSeparator()
    .addItem('Проверить Twitch ключи', 'showTwitchCredentialsStatus')
    .addToUi();
}

function showTwitchCredentialsStatus() {
  const creds = getTwitchCredentials();
  const ui = SpreadsheetApp.getUi();

  if (!creds) {
    ui.alert('❌ Twitch ключи не найдены. Добавь TWITCH_CLIENT_ID и TWITCH_CLIENT_SECRET в Script Properties (Project Settings).');
    return;
  }

  ui.alert('✅ Twitch ключи настроены.');
}

// ── АВТОЗАПОЛНЕНИЕ ССЫЛОК ПРИ ВСТАВКЕ URL ───────────────────
// Колонки: G=TwitchTracker (7), H=SCharts (8) — подправь если у тебя другие

const COL_TWITCHTRACKER = 8;  // H — TwitchTracker
const COL_SCHARTS       = 9;  // I — SCharts

function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  const range = e.range;

  // Реагируем только на колонку B (Streamer)
  if (range.getColumn() !== COL_URL) return;
  if (range.getRow() < START_ROW) return;

  const url = range.getValue();
  if (!url || typeof url !== 'string') return;

  const urlLower = url.toLowerCase();
  let username = null;
  let ttUrl = null;
  let scUrl = null;

  if (urlLower.includes('twitch.tv')) {
    username = extractUsername(url, 'twitch.tv');
    if (username) {
      ttUrl = `https://twitchtracker.com/${username}`;
      scUrl = `https://streamscharts.com/channels/${username}`;
    }
  } else if (urlLower.includes('kick.com')) {
    username = extractUsername(url, 'kick.com');
    if (username) {
      // Для Kick TwitchTracker не нужен — ставим только SCharts
      scUrl = `https://streamscharts.com/channels/${username}?platform=kick`;
    }
  }

  const row = range.getRow();

  if (ttUrl) {
    sheet.getRange(row, COL_TWITCHTRACKER).setFormula(`=HYPERLINK("${ttUrl}","click")`);
  }
  if (scUrl) {
    sheet.getRange(row, COL_SCHARTS).setFormula(`=HYPERLINK("${scUrl}","sc")`);
  }
}
