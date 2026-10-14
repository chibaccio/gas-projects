/**
 * YouTube 新着ショート動画 自動検知・解析・登録スクリプト (GAS + Gemini API)
 * 
 * 【概要】
 * 1. チャンネル「暮らしの最適化」（@lifestyle_Optimization）の最新アップロード動画を取得。
 * 2. スプレッドシート「チャンネル分析」と照合し、未登録の新規動画を特定。
 * 3. Gemini API（gemini-2.5-flash）を呼び出し、動画の画面埋め込みテロップ・文字起こしをタイムコード付きで抽出。
 * 4. Google Driveの「YouTube」フォルダにMarkdownファイル（.md）を作成・保存。
 * 5. スプレッドシート「チャンネル分析」に新規行を追加し、公開日降順でソート。
 */

const CONFIG = {
  // スプレッドシート「チャンネル分析」のID（正規化済み本番用シート）
  SPREADSHEET_ID: '1UYKSbkka7PMKzhEjZfLw9ZASlKMIbVLdEYkOxNBcN3s',
  // 保存先 Google Drive フォルダ「YouTube」のID
  YOUTUBE_FOLDER_ID: '1pUbUf_uqBZfIZuBp08fx2oi0WUmpCVj8',
  // 対象チャンネルID（@lifestyle_Optimization）
  CHANNEL_ID: 'UCsi_mLlX9ICf23WrHa4pV4Q',
  // 解析に使用するGeminiモデル
  GEMINI_MODEL: 'gemini-2.5-flash',
  // 最大取得件数（最新のアップロード動画をチェックする件数）
  MAX_CHECK_COUNT: 10
};

/**
 * 毎日定期実行するメイン関数（トリガーから呼び出し）
 */
function checkAndRegisterNewVideos() {
  console.log('=== YouTube新着動画チェック開始 ===');

  // 1. スプレッドシートの取得と既登録動画IDの取得
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheets()[0];
  const registeredIds = getRegisteredVideoIds(sheet);
  console.log(`既登録動画件数: ${registeredIds.size} 件`);

  // 2. YouTube Data API で最新アップロード動画を取得
  const latestVideos = fetchLatestVideos(CONFIG.CHANNEL_ID, CONFIG.MAX_CHECK_COUNT);
  console.log(`取得した最新動画: ${latestVideos.length} 件`);

  // 3. 未登録動画の抽出（新しい順に処理するため逆順にして登録）
  const newVideos = latestVideos.filter(v => !registeredIds.has(v.id));
  if (newVideos.length === 0) {
    console.log('未登録の新規動画はありません。終了します。');
    return;
  }
  console.log(`新規未登録動画: ${newVideos.length} 件を検出しました。`);

  // Gemini APIキーの取得（スクリプトプロパティ）
  const geminiApiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  const targetFolder = DriveApp.getFolderById(CONFIG.YOUTUBE_FOLDER_ID);

  // 公開日が古いものから順に処理して追記
  newVideos.reverse();

  for (const video of newVideos) {
    console.log(`処理開始: ${video.title} (${video.url})`);

    // 4. Gemini API による画面テロップ・文字起こし抽出
    let transcript = '';
    if (geminiApiKey) {
      try {
        transcript = generateTelopTranscriptWithGemini(video.id, geminiApiKey);
      } catch (err) {
        console.error(`Gemini API解析エラー (${video.id}): ${err.message}`);
        transcript = '- （テロップ文字起こしの自動取得に失敗しました）';
      }
    } else {
      console.warn('GEMINI_API_KEY が設定されていないため、テロップ文字起こしをスキップします。');
      transcript = '- （GEMINI_API_KEY未設定のため文字起こしなし）';
    }

    // 5. Markdownファイルの作成とDrive保存
    try {
      const mdFile = createMarkdownFile(targetFolder, video, transcript);
      console.log(`Markdown作成成功: ${mdFile.getName()} (${mdFile.getUrl()})`);
    } catch (err) {
      console.error(`Markdown作成エラー: ${err.message}`);
    }

    // 6. スプレッドシート「チャンネル分析」への追記とソート
    try {
      recordToSpreadsheet(sheet, video);
      console.log(`スプレッドシート記録成功: ${video.title}`);
    } catch (err) {
      console.error(`スプレッドシート記録エラー: ${err.message}`);
    }
  }

  // 7. 全体のソート（公開日 B列の降順）
  sortSpreadsheetByPublishDate(sheet);
  console.log('=== 新着動画チェック・登録完了 ===');
}

/**
 * スプレッドシートから登録済みの動画IDを取得
 */
function getRegisteredVideoIds(sheet) {
  const data = sheet.getDataRange().getValues();
  const registered = new Set();

  // 1行目はヘッダー、C列（インデックス2）がURL
  for (let i = 1; i < data.length; i++) {
    const url = String(data[i][2] || '').trim();
    const id = extractVideoId(url);
    if (id) registered.add(id);
  }
  return registered;
}

/**
 * URLからVideo IDを抽出
 */
function extractVideoId(url) {
  if (!url) return '';
  const match = url.match(/(?:shorts\/|v=|youtu\.be\/)([a-zA-Z0-9_-]{11})/);
  return match ? match[1] : '';
}

/**
 * YouTube Data API (Advanced Service または REST) で最新アップロード動画を取得
 * ※チャンネルのUploadsプレイリスト（UU + チャンネルID末尾）を使用し、クォータを1ユニットに最小化
 */
function fetchLatestVideos(channelId, maxResults) {
  const uploadsPlaylistId = 'UU' + channelId.substring(2);

  try {
    if (typeof YouTube !== 'undefined') {
      const response = YouTube.PlaylistItems.list('snippet,contentDetails', {
        playlistId: uploadsPlaylistId,
        maxResults: maxResults
      });
      return response.items.map(item => ({
        id: item.contentDetails.videoId,
        title: item.snippet.title,
        publishDate: Utilities.formatDate(new Date(item.snippet.publishedAt), 'Asia/Tokyo', 'yyyy/MM/dd'),
        url: `https://youtube.com/shorts/${item.contentDetails.videoId}`
      }));
    }
  } catch (e) {
    console.warn(`YouTube Advanced Service エラー: ${e.message}。OAuth経由でREST呼び出しを試行します。`);
  }

  const apiUrl = `https://www.googleapis.com/youtube/v3/playlistItems?part=snippet,contentDetails&playlistId=${uploadsPlaylistId}&maxResults=${maxResults}`;
  const response = UrlFetchApp.fetch(apiUrl, {
    headers: { Authorization: `Bearer ${ScriptApp.getOAuthToken()}` },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    throw new Error(`YouTube API呼び出し失敗 (${response.getResponseCode()}): ${response.getContentText()}`);
  }

  const data = JSON.parse(response.getContentText());
  return data.items.map(item => ({
    id: item.contentDetails.videoId,
    title: item.snippet.title,
    publishDate: Utilities.formatDate(new Date(item.snippet.publishedAt), 'Asia/Tokyo', 'yyyy/MM/dd'),
    url: `https://youtube.com/shorts/${item.contentDetails.videoId}`
  }));
}

/**
 * 動画の最新再生回数を取得 (videos.list)
 */
function fetchVideoViewCount(videoId) {
  try {
    if (typeof YouTube !== 'undefined') {
      const res = YouTube.Videos.list('statistics', { id: videoId });
      if (res.items && res.items.length > 0) {
        return Number(res.items[0].statistics.viewCount || 0);
      }
    }
  } catch (e) {
    console.warn(`Videos.list 取得警告: ${e.message}`);
  }

  const apiUrl = `https://www.googleapis.com/youtube/v3/videos?part=statistics&id=${videoId}`;
  const res = UrlFetchApp.fetch(apiUrl, {
    headers: { Authorization: `Bearer ${ScriptApp.getOAuthToken()}` },
    muteHttpExceptions: true
  });
  if (res.getResponseCode() === 200) {
    const data = JSON.parse(res.getContentText());
    if (data.items && data.items.length > 0) {
      return Number(data.items[0].statistics.viewCount || 0);
    }
  }
  return 0;
}

/**
 * Gemini API を使って画面テロップ・文字起こしをタイムコード付きで生成
 */
function generateTelopTranscriptWithGemini(videoId, apiKey) {
  const videoUrl = `https://www.youtube.com/shorts/${videoId}`;
  const prompt = `以下のYouTubeショート動画を視聴・解析し、動画内の画面埋め込みテロップおよび発話内容をタイムコード付きで文字起こししてください。

対象動画: ${videoUrl}

【出力要件】
- 以下のフォーマットに従って、タイムコード付きの箇条書きのみを出力してください。
- 挨拶や前置き、解説文、注釈は一切含めないでください。

【フォーマット例】
- 00:00 - 00:01: 名古屋名物 みそかつ 矢場とん
- 00:02 - 00:04: 名古屋といえば みそかつ
- 00:05 - 00:06: 矢場とん名物 わらじとんかつ
- 00:07 - 00:08: 行列のできる 矢場とんですが
`;

  const endpoint = `https://generativelanguage.googleapis.com/v1beta/models/${CONFIG.GEMINI_MODEL}:generateContent?key=${apiKey}`;
  const payload = {
    contents: [
      {
        parts: [
          { text: prompt }
        ]
      }
    ],
    generationConfig: {
      temperature: 0.2
    }
  };

  const response = UrlFetchApp.fetch(endpoint, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    throw new Error(`Gemini API エラー (${response.getResponseCode()}): ${response.getContentText()}`);
  }

  const result = JSON.parse(response.getContentText());
  const candidateText = result.candidates?.[0]?.content?.parts?.[0]?.text;
  if (!candidateText) {
    throw new Error('Gemini APIからのテキスト生成結果が空でした。');
  }

  return candidateText.trim();
}

/**
 * Markdownファイルを作成してDriveのYouTubeフォルダへ保存
 */
function createMarkdownFile(folder, video, transcript) {
  const views = fetchVideoViewCount(video.id);
  const formattedViews = views.toLocaleString();

  const content = `---
title: "${video.title}"
url: "${video.url}"
publish_date: "${video.publishDate}"
views: ${views}
---
# ${video.title}

## 基本情報
- **タイトル**: ${video.title}
- **URL**: ${video.url}
- **公開日**: ${video.publishDate}
- **再生回数**: ${formattedViews} 回

## 画面テロップ・文字起こし
${transcript}
`;

  const fileName = `${video.title}.md`;

  const existingFiles = folder.getFilesByName(fileName);
  if (existingFiles.hasNext()) {
    const file = existingFiles.next();
    file.setContent(content);
    return file;
  } else {
    return folder.createFile(fileName, content, MimeType.PLAIN_TEXT);
  }
}

/**
 * スプレッドシート「チャンネル分析」への新規行追加
 */
function recordToSpreadsheet(sheet, video) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const todayStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd');

  let targetColIdx = -1;
  for (let j = 3; j < headers.length; j++) {
    const colDateStr = headers[j] instanceof Date
      ? Utilities.formatDate(headers[j], 'Asia/Tokyo', 'yyyy/MM/dd')
      : String(headers[j]).trim().replace(/-/g, '/');
    if (colDateStr === todayStr) {
      targetColIdx = j + 1;
      break;
    }
  }

  if (targetColIdx === -1) {
    targetColIdx = sheet.getLastColumn() + 1;
    sheet.getRange(1, targetColIdx).setValue(todayStr);
  }

  const views = fetchVideoViewCount(video.id);
  const newRow = sheet.getLastRow() + 1;

  sheet.getRange(newRow, 1).setValue(video.title);
  sheet.getRange(newRow, 2).setValue(video.publishDate);
  sheet.getRange(newRow, 3).setValue(video.url);
  sheet.getRange(newRow, targetColIdx).setValue(views);
}

/**
 * スプレッドシートを公開日（B列）降順でソート
 */
function sortSpreadsheetByPublishDate(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow <= 1) return;

  const dataRange = sheet.getRange(2, 1, lastRow - 1, lastCol);
  dataRange.sort({ column: 2, ascending: false });
}

/**
 * テスト用ドライラン（書き込みを行わずシート照合と最新動画取得を検証）
 */
function testDryRun() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheets()[0];
  const registeredIds = getRegisteredVideoIds(sheet);
  const latestVideos = fetchLatestVideos(CONFIG.CHANNEL_ID, CONFIG.MAX_CHECK_COUNT);
  const newVideos = latestVideos.filter(v => !registeredIds.has(v.id));
  console.log(`DryRun完了 - 既登録: ${registeredIds.size}件, 取得: ${latestVideos.length}件, 未登録: ${newVideos.length}件`);
  return { registeredCount: registeredIds.size, fetchedCount: latestVideos.length, newCount: newVideos.length };
}

/**
 * 【設定用】毎日定時（例: 15:00）に自動実行するトリガーを登録
 */
function setDailyTrigger() {
  const functionName = 'checkAndRegisterNewVideos';
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === functionName) {
      ScriptApp.deleteTrigger(trigger);
    }
  }

  ScriptApp.newTrigger(functionName)
    .timeBased()
    .atHour(15) // 15:00〜16:00に実行
    .everyDays(1)
    .inTimezone('Asia/Tokyo')
    .create();

  console.log('毎日15:00〜16:00の定期実行トリガーを登録しました。');
}
