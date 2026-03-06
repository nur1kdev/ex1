// ============================================================
//  АВТООБНОВЛЕНИЕ STREAMERS  v7.0 FINAL (WORKING)
//  Twitch: AVG viewers + Followers ✅
//  Kick: AVG viewers ✅ + Followers ✅ (NEW API)
// ============================================================

// ── ТВОИ КЛЮЧИ TWITCH ───────────────────────────────────────
const SCRIPT_PROP_TWITCH_CLIENT_ID = 'TWITCH_CLIENT_ID';
const SCRIPT_PROP_TWITCH_CLIENT_SECRET = 'TWITCH_CLIENT_SECRET';
const TWITCH_CLIENT_ID = ''; // optional fallback (если не хотите Script Properties)
const TWITCH_CLIENT_SECRET = ''; // optional fallback
// ───────────────────────────────────────────────────────────

// ── НАСТРОЙКИ КОЛОНОК ───────────────────────────────────────
const COL_URL = 2; // B — ссылка на стримера
const COL_AVG_ONLINE = 5; // E — AVG Online
const COL_FOLLOWERS = 6; // F — Followers
const START_ROW = 2;
const COL_TWITCHTRACKER = 8; // H — TwitchTracker
const COL_SCHARTS = 9; // I — SCharts
// ───────────────────────────────────────────────────────────

// ── TWITCH CREDENTIALS (Script Properties) ─────────────────

function getTwitchCredentials() {
  const props = PropertiesService.getScriptProperties();
  const propClientId = props.getProperty(SCRIPT_PROP_TWITCH_CLIENT_ID);
  const propClientSecret = props.getProperty(SCRIPT_PROP_TWITCH_CLIENT_SECRET);

  const clientId = propClientId || TWITCH_CLIENT_ID;
  const clientSecret = propClientSecret || TWITCH_CLIENT_SECRET;

  if (!clientId || !clientSecret) {
    return null;
  }

  return { clientId, clientSecret };
}

function setTwitchCredentials(clientId, clientSecret) {
  if (!clientId || !clientSecret) {
    throw new Error('Передайте TWITCH_CLIENT_ID и TWITCH_CLIENT_SECRET');
  }

  const props = PropertiesService.getScriptProperties();
  props.setProperty(SCRIPT_PROP_TWITCH_CLIENT_ID, clientId);
  props.setProperty(SCRIPT_PROP_TWITCH_CLIENT_SECRET, clientSecret);
}

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

  let token = null;

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

        if (!token) {
          token = getTwitchToken();
          if (!token) {
            Logger.log('⚠️ Twitch token недоступен, используем fallback для followers без Helix');
          }
        }

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

    if (!token) {
      Logger.log('⚠️ Twitch token недоступен, используем fallback для followers без Helix');
    }

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

  const creds = getTwitchCredentials();
  if (!creds) return null;

  try {
    const resp = UrlFetchApp.fetch('https://id.twitch.tv/oauth2/token', {
      method: 'post',
      payload: {
        client_id: creds.clientId,
        client_secret: creds.clientSecret,
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
  const creds = getTwitchCredentials();

  // 1) Основной путь: Twitch Helix
  if (creds && token) {
    try {
      const userResp = UrlFetchApp.fetch(`https://api.twitch.tv/helix/users?login=${encodeURIComponent(username)}`, {
        headers: { 'Client-ID': creds.clientId, Authorization: `Bearer ${token}` },
        muteHttpExceptions: true,
      });

      const userData = JSON.parse(userResp.getContentText());
      if (userData?.data?.length) {
        const userId = userData.data[0].id;

        const fResp = UrlFetchApp.fetch(`https://api.twitch.tv/helix/channels/followers?broadcaster_id=${userId}&first=1`, {
          headers: { 'Client-ID': creds.clientId, Authorization: `Bearer ${token}` },
          muteHttpExceptions: true,
        });

        if (fResp.getResponseCode() === 200) {
          const fData = JSON.parse(fResp.getContentText());
          if (typeof fData?.total === 'number') return fData.total;
        }
      }
    } catch (e) {
      Logger.log(`Twitch Helix followers ошибка ${username}: ${e.message}`);
    }
  }

  // 2) Fallback без ключей: ivr.fi
  try {
    const ivrResp = UrlFetchApp.fetch(`https://api.ivr.fi/v2/twitch/user?login=${encodeURIComponent(username)}`, {
      method: 'get',
      headers: { 'User-Agent': 'Mozilla/5.0', Accept: 'application/json' },
      muteHttpExceptions: true,
      followRedirects: true,
    });

    if (ivrResp.getResponseCode() !== 200) return null;

    const ivrData = JSON.parse(ivrResp.getContentText());
    const user = Array.isArray(ivrData) ? ivrData[0] : ivrData;
    if (!user) return null;

    return user.followers ?? user.followersCount ?? null;
  } catch (e) {
    Logger.log(`IVR followers ошибка ${username}: ${e.message}`);
    return null;
  }
}

// ── KICK — AVG VIEWERS (AeroKick) ───────────────────────────
function fetchAeroKickAvg(username) {
  const normalized = username.toLowerCase();
  const aeroKickUrl = `https://aerokick.app/stats/channels/${normalized}?range=month`;

  try {
    const resp = UrlFetchApp.fetch(aeroKickUrl, {
      method: 'get',
      headers: { 'User-Agent': 'Mozilla/5.0', Accept: 'text/html,application/json' },
      muteHttpExceptions: true,
      followRedirects: true,
    });

    if (resp.getResponseCode() === 200) {
      const body = resp.getContentText();
      const parsed = tryParseAvgViewers(body);
      if (parsed !== null) return parsed;
    }
  } catch (e) {
    Logger.log(`AeroKick AVG ошибка ${username}: ${e.message}`);
  }

  // Fallback: Streamscharts (Kick)
  try {
    const scUrl = `https://streamscharts.com/channels/${normalized}?platform=kick`;
    const resp = UrlFetchApp.fetch(scUrl, {
      method: 'get',
      headers: { 'User-Agent': 'Mozilla/5.0', Accept: 'text/html' },
      muteHttpExceptions: true,
      followRedirects: true,
    });

    if (resp.getResponseCode() !== 200) return null;

    const body = resp.getContentText();
    return tryParseAvgViewers(body);
  } catch (e) {
    Logger.log(`Streamscharts AVG ошибка ${username}: ${e.message}`);
    return null;
  }
}

// ── KICK — FOLLOWERS ✅ NEW API (WORKING) ───────────────────
function parseFlexibleNumber(rawValue) {
  if (rawValue === null || rawValue === undefined) return null;

  const cleaned = String(rawValue).trim().replace(/\s+/g, '');
  if (!cleaned) return null;

  // Если есть и точка, и запятая — убираем запятые как разделители тысяч
  if (cleaned.includes('.') && cleaned.includes(',')) {
    const n = Number(cleaned.replace(/,/g, ''));
    return Number.isFinite(n) ? n : null;
  }

  // Если только запятая: может быть 1,234 (тысячи) или 12,3 (десятичные)
  if (cleaned.includes(',') && !cleaned.includes('.')) {
    const parts = cleaned.split(',');
    const normalized = parts[parts.length - 1].length === 3 ? cleaned.replace(/,/g, '') : cleaned.replace(',', '.');
    const n = Number(normalized);
    return Number.isFinite(n) ? n : null;
  }

  const n = Number(cleaned);
  return Number.isFinite(n) ? n : null;
}

function tryParseAvgViewers(text) {
  if (!text) return null;

  const patterns = [
    /"avg[_-]?viewers"\s*:\s*"?([\d\s.,]+)"?/i,
    /"average[_-]?viewers"\s*:\s*"?([\d\s.,]+)"?/i,
    /"avgViewers"\s*:\s*"?([\d\s.,]+)"?/i,
    /"value",\s*([\d\s.,]+)\]/i,
    /Average Viewers[^\d]{0,30}([\d\s.,]+)/i,
  ];

  for (const pattern of patterns) {
    const match = text.match(pattern);
    if (!match) continue;

    const value = parseFlexibleNumber(match[1]);
    if (Number.isFinite(value)) return Math.round(value);
  }

  return null;
}

function fetchKickFollowers(username) {
  const normalized = username.toLowerCase();
  const urls = [
    `https://api.kick.com/private/v1/channels/${normalized}`,
    `https://kick.com/api/v2/channels/${normalized}`,
    `https://api.kick.com/public/v1/channels/${normalized}`,
  ];

  for (const url of urls) {
    try {
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
        Logger.log(`Kick API HTTP ${resp.getResponseCode()} (${url}) для ${username}`);
        continue;
      }

      const data = JSON.parse(resp.getContentText());
      const followers =
        data?.data?.channel?.followers_count ??
        data?.channel?.followers_count ??
        data?.followers_count ??
        data?.data?.followers_count ??
        null;

      if (followers !== null) {
        Logger.log(`✅ Kick followers для ${username}: ${followers}`);
        return Number(followers);
      }
    } catch (e) {
      Logger.log(`Ошибка Kick API ${username} (${url}): ${e.message}`);
    }
  }

  Logger.log(`❌ Не найдены фолловеры для ${username}`);
  return null;
}

// ── УТИЛИТЫ ──────────────────────────────────────────────────
function extractUsername(url, domain) {
  const regex = new RegExp(domain.replace('.', '\\.') + '\\/([a-zA-Z0-9_.-]+)');
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
    .addSeparator()
    .addItem('🔐 Проверить Twitch ключи', 'showTwitchCredentialsStatus')
    .addToUi();
}

function showTwitchCredentialsStatus() {
  const ui = SpreadsheetApp.getUi();
  const creds = getTwitchCredentials();

  if (!creds) {
    ui.alert('❌ Twitch ключи не настроены в Script Properties.\n\nОткройте: Project Settings → Script properties и добавьте TWITCH_CLIENT_ID / TWITCH_CLIENT_SECRET.');
    return;
  }

  ui.alert('✅ Twitch ключи найдены в Script Properties.');
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
