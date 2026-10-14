/**
 * YouTube Analytics Tracker (GAS)
 * 
 * 目的:
 * YouTubeチャンネル「暮らしの最適化」（@lifestyle_Optimization）の
 * 1. チャンネル登録者数・総動画本数の日次追記
 * 2. 各動画の最新再生回数の日次追記（日次列追加）
 * 3. 新規公開動画の自動検知・行追加
 * をGoogle Apps Script単体で完全自律稼働させるスクリプト。
 * 
 * 【事前準備】
 * GASエディタの左メニュー「サービス (+)」から「YouTube Data API v3」を追加してください。
 */

const CONFIG = {
  // 対象YouTubeチャンネル情報
  CHANNEL_ID: 'UCsi_mLlX9ICf23WrHa4pV4Q', // 暮らしの最適化 (@lifestyle_Optimization)

  // 対象スプレッドシートID（正規化済み本番用シート）
  SPREADSHEET_ID_SUBSCRIBERS: '1JXUH-r9NEuJ8Q8ZAj_oMWP-q7baIe7ddvlTshvuBykU', // YouTubeチャンネル登録者数
  SPREADSHEET_ID_VIEWS: '1UYKSbkka7PMKzhEjZfLw9ZASlKMIbVLdEYkOxNBcN3s',       // チャンネル分析

  TIMEZONE: 'Asia/Tokyo',
  DATE_FORMAT: 'yyyy/MM/dd'
};

/**
 * 毎日定時（午前3:00〜4:00）に実行するメイン関数
 */
function runDailyTracking() {
  console.log('=== YouTube 日次トラッキング開始 ===');
  const now = new Date();
  const todayStr = Utilities.formatDate(now, CONFIG.TIMEZONE, CONFIG.DATE_FORMAT);
  console.log(`実行日時: ${todayStr}`);

  try {
    // 1. チャンネル登録者数・動画本数の記録
    recordChannelSubscribers(todayStr);
  } catch (e) {
    console.error(`チャンネル登録者数の記録でエラーが発生しました: ${e.message}`);
  }

  try {
    // 2. 各動画の最新再生回数の記録（新規動画の検知・追加含む）
    recordVideoViews(todayStr);
  } catch (e) {
    console.error(`動画再生回数の記録でエラーが発生しました: ${e.message}`);
  }

  console.log('=== YouTube 日次トラッキング完了 ===');
}

/**
 * 1. チャンネル登録者数・動画本数の追記処理
 * @param {string} todayStr 'YYYY/MM/DD'
 */
function recordChannelSubscribers(todayStr) {
  console.log('--- 登録者数・動画本数の取得・記録 ---');

  const channelRes = YouTube.Channels.list('statistics', {
    id: CONFIG.CHANNEL_ID
  });

  if (!channelRes.items || channelRes.items.length === 0) {
    throw new Error(`チャンネル情報が見つかりませんでした: ID=${CONFIG.CHANNEL_ID}`);
  }

  const stats = channelRes.items[0].statistics;
  const subscriberCount = Number(stats.subscriberCount);
  const videoCount = Number(stats.videoCount);
  console.log(`取得結果 - 登録者数: ${subscriberCount}, 動画本数: ${videoCount}`);

  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID_SUBSCRIBERS);
  const sheet = ss.getSheets()[0];
  const lastRow = sheet.getLastRow();

  // 重複チェック（最終行が今日の日付の場合は上書き更新、異なる場合は行追加）
  if (lastRow > 1) {
    const lastRowDateVal = sheet.getRange(lastRow, 1).getValue();
    const lastRowDateStr = lastRowDateVal instanceof Date 
      ? Utilities.formatDate(lastRowDateVal, CONFIG.TIMEZONE, CONFIG.DATE_FORMAT)
      : String(lastRowDateVal).replace(/-/g, '/');

    if (lastRowDateStr === todayStr) {
      console.log(`本日（${todayStr}）のデータが既に最終行（${lastRow}行目）に存在するため上書き更新します。`);
      sheet.getRange(lastRow, 2, 1, 2).setValues([[subscriberCount, videoCount]]);
      return;
    }
  }

  sheet.appendRow([todayStr, subscriberCount, videoCount]);
  console.log(`最下行（${lastRow + 1}行目）に [${todayStr}, ${subscriberCount}, ${videoCount}] を追記しました。`);
}

/**
 * 2. 各動画の再生回数の日次追記処理（新規動画の検知・追加含む）
 * @param {string} todayStr 'YYYY/MM/DD'
 */
function recordVideoViews(todayStr) {
  console.log('--- 動画再生回数の取得・記録 ---');

  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID_VIEWS);
  const sheet = ss.getSheets()[0];

  const data = sheet.getDataRange().getValues();
  if (data.length < 1) {
    throw new Error('シートにデータが存在しません。');
  }

  const headerRow = data[0];
  let lastCol = sheet.getLastColumn();

  let targetColIndex = -1;
  for (let c = 3; c < headerRow.length; c++) {
    const colVal = headerRow[c];
    const colStr = colVal instanceof Date
      ? Utilities.formatDate(colVal, CONFIG.TIMEZONE, CONFIG.DATE_FORMAT)
      : String(colVal).replace(/-/g, '/');
    if (colStr === todayStr) {
      targetColIndex = c + 1;
      break;
    }
  }

  if (targetColIndex === -1) {
    targetColIndex = lastCol + 1;
    sheet.getRange(1, targetColIndex).setValue(todayStr);
    console.log(`新しく本日の日付列（${targetColIndex}列目: ${todayStr}）を追加しました。`);
  } else {
    console.log(`本日の日付列は既に ${targetColIndex}列目に存在します。値を更新します。`);
  }

  const existingVideoMap = new Map();
  const videoIdsToFetch = [];

  for (let r = 1; r < data.length; r++) {
    const url = String(data[r][2] || '').trim();
    const videoId = extractVideoId(url);
    if (videoId) {
      existingVideoMap.set(videoId, r);
      videoIdsToFetch.push(videoId);
    }
  }

  console.log(`シート内の追跡対象動画数: ${videoIdsToFetch.length}本`);

  const viewsMap = fetchVideoViewsBatch(videoIdsToFetch);

  const updateRowsCount = data.length - 1;
  if (updateRowsCount > 0) {
    const valuesColumn = [];
    for (let r = 1; r < data.length; r++) {
      const url = String(data[r][2] || '').trim();
      const videoId = extractVideoId(url);
      const views = (videoId && viewsMap.has(videoId)) ? viewsMap.get(videoId) : '';
      valuesColumn.push([views]);
    }

    sheet.getRange(2, targetColIndex, valuesColumn.length, 1).setValues(valuesColumn);
    console.log(`既存 ${valuesColumn.length}本の動画の再生回数を反映しました。`);
  }

  try {
    checkAndAddNewVideos(sheet, existingVideoMap, targetColIndex);
  } catch (err) {
    console.warn(`新規動画の自動検出で警告（スキップ）: ${err.message}`);
  }
}

/**
 * 動画URLからYouTubeの動画ID（11桁）を抽出
 */
function extractVideoId(url) {
  if (!url) return null;
  const match = url.match(/(?:shorts\/|v=|youtu\.be\/|\/v\/|\/embed\/)([a-zA-Z0-9_-]{11})/);
  return match ? match[1] : null;
}

/**
 * 動画ID配列からYouTube Data APIを使って一括で再生回数を取得（最大50件ずつ）
 */
function fetchVideoViewsBatch(videoIds) {
  const viewsMap = new Map();
  const chunkSize = 50;

  for (let i = 0; i < videoIds.length; i += chunkSize) {
    const chunk = videoIds.slice(i, i + chunkSize);
    const res = YouTube.Videos.list('statistics', {
      id: chunk.join(',')
    });

    if (res.items) {
      for (const item of res.items) {
        const vId = item.id;
        const count = Number(item.statistics.viewCount || 0);
        viewsMap.set(vId, count);
      }
    }
  }

  return viewsMap;
}

/**
 * チャンネルの最新公開動画（最大15件）をチェックし、シートに未登録の動画があれば自動で新規行を追加
 */
function checkAndAddNewVideos(sheet, existingVideoMap, todayColIndex) {
  const uploadsPlaylistId = CONFIG.CHANNEL_ID.replace(/^UC/, 'UU');

  const res = YouTube.PlaylistItems.list('snippet,contentDetails', {
    playlistId: uploadsPlaylistId,
    maxResults: 15
  });

  if (!res.items || res.items.length === 0) return;

  const newVideosToAdd = [];
  for (const item of res.items) {
    const videoId = item.contentDetails.videoId;
    if (!existingVideoMap.has(videoId)) {
      const title = item.snippet.title;
      const publishedAt = item.snippet.publishedAt ? item.snippet.publishedAt.split('T')[0].replace(/-/g, '/') : '';
      const url = `https://youtube.com/shorts/${videoId}`;
      newVideosToAdd.push({ videoId, title, publishedAt, url });
    }
  }

  if (newVideosToAdd.length > 0) {
    console.log(`未登録の新規動画を ${newVideosToAdd.length}件 検出しました。行を追加します。`);
    const newViewsMap = fetchVideoViewsBatch(newVideosToAdd.map(v => v.videoId));
    const totalCols = sheet.getLastColumn();

    for (const v of newVideosToAdd) {
      const newRow = new Array(totalCols).fill('');
      newRow[0] = v.title;
      newRow[1] = v.publishedAt;
      newRow[2] = v.url;
      const views = newViewsMap.get(v.videoId) || 0;
      newRow[todayColIndex - 1] = views;
      sheet.appendRow(newRow);
      console.log(`新規動画を追加: ${v.title} (${v.url}) - 再生回数: ${views}`);
    }
  }
}

/**
 * 毎日午前3:00〜4:00に自動実行する時間主導型トリガーを登録・再設定する関数
 */
function setDailyTrigger() {
  const functionName = 'runDailyTracking';
  const triggers = ScriptApp.getProjectTriggers();

  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === functionName) {
      ScriptApp.deleteTrigger(trigger);
    }
  }

  ScriptApp.newTrigger(functionName)
    .timeBased()
    .atHour(3)
    .everyDays(1)
    .inTimezone(CONFIG.TIMEZONE)
    .create();

  console.log(`毎日午前3時台に自動実行するトリガーを登録しました（関数: ${functionName}）。`);
}

/**
 * テスト用: チャンネル登録者数・動画本数のみ記録
 */
function testSubscribersOnly() {
  const todayStr = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, CONFIG.DATE_FORMAT);
  recordChannelSubscribers(todayStr);
}

/**
 * テスト用: 動画再生回数のみ記録
 */
function testViewsOnly() {
  const todayStr = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, CONFIG.DATE_FORMAT);
  recordVideoViews(todayStr);
}
