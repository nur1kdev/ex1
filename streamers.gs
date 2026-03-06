// ============================================================
//  АВТООБНОВЛЕНИЕ STREAMERS  v7.0 FINAL (WORKING)
//  Twitch: AVG viewers + Followers ✅
//  Kick: AVG viewers ✅ + Followers ✅ (NEW API)
// ============================================================

// ── ТВОИ КЛЮЧИ TWITCH ───────────────────────────────────────
const TWITCH_CLIENT_ID = 'e961qn1qk18bk8gk5hcmz61c687y8i';
const TWITCH_CLIENT_SECRET = 'pei0rmegkkmjrklfvlo6mq10d857ff';
// ───────────────────────────────────────────────────────────

// ── НАСТРОЙКИ КОЛОНОК ───────────────────────────────────────
const COL_URL = 2; // B — ссылка на стримера
const COL_AVG_ONLINE = 5; // E — AVG Online
const COL_FOLLOWERS = 6; // F — Followers
const START_ROW = 2;
const COL_TWITCHTRACKER = 8; // H — TwitchTracker
const COL_SCHARTS = 9; // I — SCharts
// ───────────────────────────────────────────────────────────

// ── ОБНОВИТЬ ВСЕХ ───────────────────────────────────────────
function updateAllStreamers() {
  updateStreamersByPlatform('all');
}

// ── ОБНОВИТЬ ТОЛЬКО TWITCH ──────────────────────────────────
function updateTwitchStreamers() {
  updateStreamersByPlatform('twitch');
}

// ── ОБНОВИТЬ ТОЛЬКО KICK ────────────────────────────────────
function updateKickStreamers() {
  updateStreamersByPlatform('kick');
}

// ── ОБЩАЯ ФУНКЦИЯ ОБНОВЛЕНИЯ ────────────────────────────────
function updateStreamersByPlatform(platform) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const sheetName = sheet.getName();
  const lastRow = sheet.getLastRow();

  const platformText = platform === 'twitch' ? '🟣 Twitch' : platform === 'kick' ? '🟢 Kick' : '🔄 Все';

  if (lastRow < START_ROW) {
    SpreadsheetApp.getUi().alert('❌ Нет данных на листе "' + sheetName + '"');
    return;
  }

  const token = getTwitchToken();
  if (!token && platform !== 'kick') {
    SpreadsheetApp.getUi().alert('❌ Ошибка Twitch токена');
    return;
  }

  let updated = 0;
  let skipped = 0;

  for (let row = START_ROW; row <= lastRow; row++) {
    const url = sheet.getRange(row, COL_URL).getValue();
    if (!url || typeof url !== 'string') continue;

    const urlLower = url.toLowerCase().trim();
    const isTwitch = urlLower.includes('twitch.tv');
    const isKick = urlLower.includes('kick.com');

    if (platform === 'twitch' && !isTwitch) {
      skipped++;
      continue;
    }
    if (platform === 'kick' && !isKick) {
      skipped++;
      continue;
    }
    if (platform === 'all' && !isTwitch && !isKick) {
      skipped++;
      continue;
    }

    try {
      if (isTwitch) {
        const username = extractUsername(url, 'twitch.tv');
        if (!username) continue;

        const avgOnline = fetchTwitchTrackerAvg(username);
        const followers = fetchTwitchFollowers(username, token);

        if (avgOnline !== null) sheet.getRange(row, COL_AVG_ONLINE).setValue(avgOnline);
        if (followers !== null) sheet.getRange(row, COL_FOLLOWERS).setValue(followers);
        updated++;
        Utilities.sleep(1500);
      } else if (isKick) {
        const username = extractUsername(url, 'kick.com');
        if (!username) continue;

        const avgOnline = fetchAeroKickAvg(username);
        const followers = fetchKickFollowers(username); // ✅ NEW API

        if (avgOnline !== null) sheet.getRange(row, COL_AVG_ONLINE).setValue(avgOnline);
        if (followers !== null) sheet.getRange(row, COL_FOLLOWERS).setValue(followers);
        updated++;
        Utilities.sleep(800);
      }
    } catch (e) {
      Logger.log(`Ошибка строка ${row}: ${e.message}`);
    }
  }

  let msg = `📄 ${sheetName}\n🎮 ${platformText}\n✅ Обновлено: ${updated}`;
  if (skipped > 0) msg += `\n⏭️ Пропущено: ${skipped}`;
  SpreadsheetApp.getUi().alert(msg);
}

// ── ОБНОВИТЬ ОДНУ СТРОКУ ─────────────────────────────────────
function updateSelectedRow() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const row = sheet.getActiveRange().getRow();

  const url = sheet.getRange(row, COL_URL).getValue();
  if (!url) {
    ui.alert(`Нет ссылки в строке ${row}`);
    return;
  }

  if (ui.alert('Обновить?', url, ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;

  const urlLower = url.toLowerCase().trim();
  let avgOnline = null;
  let followers = null;

  if (urlLower.includes('twitch.tv')) {
    const username = extractUsername(url, 'twitch.tv');
    const token = getTwitchToken();
    avgOnline = fetchTwitchTrackerAvg(username);
    followers = fetchTwitchFollowers(username, token);
  } else if (urlLower.includes('kick.com')) {
    const username = extractUsername(url, 'kick.com');
    avgOnline = fetchAeroKickAvg(username);
    followers = fetchKickFollowers(username);
  }

  if (avgOnline !== null) sheet.getRange(row, COL_AVG_ONLINE).setValue(avgOnline);
  if (followers !== null) sheet.getRange(row, COL_FOLLOWERS).setValue(followers);

  ui.alert(`✅ Готово!\nAVG: ${avgOnline ?? '-'}\nFollowers: ${followers ?? '-'}`);
}

// ── TWITCHTRACKER API — AVG VIEWERS ─────────────────────────
function fetchTwitchTrackerAvg(username) {
  const apiUrl = `https://twitchtracker.com/api/channels/summary/${username.toLowerCase()}`;

  let cookieHeader = '';
  try {
    const pageResp = UrlFetchApp.fetch(`https://twitchtracker.com/${username.toLowerCase()}`, {
      method: 'get',
      headers: { 'User-Agent': 'Mozilla/5.0' },
      muteHttpExceptions: true,
    });
    const setCookie = pageResp.getAllHeaders()['Set-Cookie'];
    if (setCookie) {
      const cookieArr = Array.isArray(setCookie) ? setCookie : [setCookie];
      cookieHeader = cookieArr.map((c) => c.split(';')[0]).join('; ');
    }
  } catch (e) {
    // noop
  }

  try {
    const resp = UrlFetchApp.fetch(apiUrl, {
      method: 'get',
      headers: {
        'User-Agent': 'Mozilla/5.0',
        Accept: 'application/json',
        Cookie: cookieHeader,
      },
      muteHttpExceptions: true,
    });

    if (resp.getResponseCode() !== 200) return null;

    const data = JSON.parse(resp.getContentText());
    return data.avg_viewers ?? data.avg_cviewers ?? null;
  } catch (e) {
    return null;
  }
}

// ── TWITCH API — TOKEN & FOLLOWERS ───────────────────────────
function getTwitchToken() {
  const cache = PropertiesService.getScriptProperties();
  const saved = cache.getProperty('twitch_token');
  const exp = cache.getProperty('twitch_token_exp');

  if (saved && exp && Date.now() < parseInt(exp, 10)) return saved;

  try {
    const resp = UrlFetchApp.fetch('https://id.twitch.tv/oauth2/token', {
      method: 'post',
      payload: {
        client_id: TWITCH_CLIENT_ID,
        client_secret: TWITCH_CLIENT_SECRET,
        grant_type: 'client_credentials',
      },
      muteHttpExceptions: true,
    });

    const data = JSON.parse(resp.getContentText());
    if (!data.access_token) return null;

    cache.setProperty('twitch_token', data.access_token);
    cache.setProperty('twitch_token_exp', (Date.now() + 30 * 24 * 3600 * 1000).toString());
    return data.access_token;
  } catch (e) {
    return null;
  }
}

function fetchTwitchFollowers(username, token) {
  try {
    const userResp = UrlFetchApp.fetch(`https://api.twitch.tv/helix/users?login=${encodeURIComponent(username)}`, {
      headers: { 'Client-ID': TWITCH_CLIENT_ID, Authorization: `Bearer ${token}` },
      muteHttpExceptions: true,
    });

    const userData = JSON.parse(userResp.getContentText());
    if (!userData.data || userData.data.length === 0) return null;

    const userId = userData.data[0].id;

    const fResp = UrlFetchApp.fetch(`https://api.twitch.tv/helix/channels/followers?broadcaster_id=${userId}&first=1`, {
      headers: { 'Client-ID': TWITCH_CLIENT_ID, Authorization: `Bearer ${token}` },
      muteHttpExceptions: true,
    });

    const fData = JSON.parse(fResp.getContentText());
    return fData.total ?? null;
  } catch (e) {
    return null;
  }
}

// ── KICK — AVG VIEWERS (AeroKick) ───────────────────────────
function fetchAeroKickAvg(username) {
  const pageUrl = `https://aerokick.app/stats/channels/${username.toLowerCase()}?range=month`;

  try {
    const resp = UrlFetchApp.fetch(pageUrl, {
      method: 'get',
      headers: { 'User-Agent': 'Mozilla/5.0', Accept: 'text/html' },
      muteHttpExceptions: true,
    });

    if (resp.getResponseCode() !== 200) return null;

    const body = resp.getContentText();
    const valueMatch = body.match(/"value",(\d+\.?\d*)\]/);

    if (valueMatch) {
      return Math.round(parseFloat(valueMatch[1]));
    }

    return null;
  } catch (e) {
    return null;
  }
}

// ── KICK — FOLLOWERS ✅ NEW API (WORKING) ───────────────────
function fetchKickFollowers(username) {
  try {
    // ✅ НОВЫЙ ENDPOINT который работает!
    const url = `https://api.kick.com/private/v1/channels/${username}`;

    const resp = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: {
        Accept: 'application/json',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
        Origin: 'https://kick.com',
        Referer: 'https://kick.com/',
      },
      muteHttpExceptions: true,
      followRedirects: true,
    });

    if (resp.getResponseCode() !== 200) {
      Logger.log(`Kick API HTTP ${resp.getResponseCode()} для ${username}`);
      return null;
    }

    const data = JSON.parse(resp.getContentText());

    // ✅ Извлекаем из data.data.channel.followers_count
    const followers = data?.data?.channel?.followers_count ?? null;

    if (followers !== null) {
      Logger.log(`✅ Kick followers для ${username}: ${followers}`);
      return followers;
    }

    Logger.log(`❌ Не найдены фолловеры для ${username}`);
    return null;
  } catch (e) {
    Logger.log(`Ошибка Kick API ${username}: ${e.message}`);
    return null;
  }
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
    .addItem('🔄 Обновить всех', 'updateAllStreamers')
    .addSeparator()
    .addItem('🟣 Обновить Twitch', 'updateTwitchStreamers')
    .addItem('🟢 Обновить Kick', 'updateKickStreamers')
    .addSeparator()
    .addItem('🎯 Обновить строку', 'updateSelectedRow')
    .addToUi();
}

// ── АВТОЗАПОЛНЕНИЕ ССЫЛОК ПРИ ВСТАВКЕ URL ───────────────────
function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  const range = e.range;

  if (range.getColumn() !== COL_URL || range.getRow() < START_ROW) return;

  const url = range.getValue();
  if (!url || typeof url !== 'string') return;

  const urlLower = url.toLowerCase().trim();
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
      scUrl = `https://streamscharts.com/channels/${username}?platform=kick`;
    }
  }

  const row = range.getRow();
  if (ttUrl) sheet.getRange(row, COL_TWITCHTRACKER).setFormula(`=HYPERLINK("${ttUrl}","click")`);
  if (scUrl) sheet.getRange(row, COL_SCHARTS).setFormula(`=HYPERLINK("${scUrl}","sc")`);
}
