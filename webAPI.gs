// 設定を関数内で取得するように変更
function getSheetId() {
  return PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
}

function getApiKey() {
  return PropertiesService.getScriptProperties().getProperty('API_KEY'); // オプション：APIキー認証用
}

// メインのGETリクエスト処理
function doGet(e) {
  try {
    // APIキー認証（必要な場合）
    const apiKey = getApiKey();
    if (apiKey && e.parameter.apiKey !== apiKey) {
      return createErrorResponse('Unauthorized', 401);
    }

    const action = e.parameter.action || 'list';
    
    switch (action) {
      case 'list':
        return handleListEvents(e);
      case 'search':
        return handleSearchEvents(e);
      case 'categories':
        return handleGetCategories(e);
      case 'byCategory':
        return handleGetEventsByCategory(e);
      case 'upcoming':
        return handleGetUpcomingEvents(e);
      default:
        return createErrorResponse('Invalid action', 400);
    }
    
  } catch (error) {
    console.error('Error:', error);
    return createErrorResponse(error.toString(), 500);
  }
}

// イベント一覧を取得
function handleListEvents(e) {
  const sheetName = e.parameter.sheet || 'イベント一覧';
  const limit = parseInt(e.parameter.limit) || 100;
  const offset = parseInt(e.parameter.offset) || 0;
  
  const events = getEventsFromSheet(sheetName, limit, offset);
  
  return createSuccessResponse({
    events: events,
    total: getTotalEventCount(sheetName),
    limit: limit,
    offset: offset
  });
}

// イベント検索
function handleSearchEvents(e) {
  const query = e.parameter.q || '';
  const sheetName = e.parameter.sheet || 'イベント一覧';
  
  if (!query) {
    return createErrorResponse('Search query is required', 400);
  }
  
  const allEvents = getEventsFromSheet(sheetName, 1000, 0);
  const filteredEvents = allEvents.filter(event => {
    const searchText = `${event.title} ${event.summary} ${event.category}`.toLowerCase();
    return searchText.includes(query.toLowerCase());
  });
  
  return createSuccessResponse({
    events: filteredEvents,
    query: query,
    count: filteredEvents.length
  });
}

// カテゴリ一覧を取得
function handleGetCategories(e) {
  const sheetId = getSheetId();
  if (!sheetId) {
    return createErrorResponse('SPREADSHEET_ID is not set', 500);
  }
  
  const ss = SpreadsheetApp.openById(sheetId);
  const sheets = ss.getSheets();
  const categories = ['課外活動', '大学進学', '語学・留学', '重要', 'その他'];
  
  const categoryInfo = categories.map(cat => {
    const sheet = ss.getSheetByName(cat);
    return {
      name: cat,
      count: sheet ? sheet.getLastRow() : 0
    };
  });
  
  return createSuccessResponse({
    categories: categoryInfo
  });
}

// カテゴリ別イベント取得
function handleGetEventsByCategory(e) {
  const category = e.parameter.category;
  
  if (!category) {
    return createErrorResponse('Category is required', 400);
  }
  
  const events = getEventsFromSheet(category, 100, 0);
  
  return createSuccessResponse({
    category: category,
    events: events,
    count: events.length
  });
}

// 今後のイベント取得（締切が近い順）
function handleGetUpcomingEvents(e) {
  const days = parseInt(e.parameter.days) || 30;
  const sheetName = e.parameter.sheet || 'イベント一覧';
  
  const today = new Date();
  const futureDate = new Date();
  futureDate.setDate(today.getDate() + days);
  
  const allEvents = getEventsFromSheet(sheetName, 1000, 0);
  const upcomingEvents = allEvents.filter(event => {
    if (!event.deadline) return false;
    const eventDate = new Date(event.deadline);
    return eventDate >= today && eventDate <= futureDate;
  }).sort((a, b) => new Date(a.deadline) - new Date(b.deadline));
  
  return createSuccessResponse({
    events: upcomingEvents,
    fromDate: today.toISOString().split('T')[0],
    toDate: futureDate.toISOString().split('T')[0],
    count: upcomingEvents.length
  });
}

// シートからイベントデータを取得
function getEventsFromSheet(sheetName, limit, offset) {
  const sheetId = getSheetId();
  if (!sheetId) {
    throw new Error('SPREADSHEET_ID is not set in Script Properties');
  }
  
  const ss = SpreadsheetApp.openById(sheetId);
  const sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    return [];
  }
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 1) return [];
  
  const startRow = Math.min(offset + 1, lastRow);
  const numRows = Math.min(limit, lastRow - offset);
  
  if (numRows <= 0) return [];
  
  const range = sheet.getRange(startRow, 1, numRows, 5);
  const values = range.getValues();
  
  return values.map((row, index) => ({
    id: offset + index + 1,
    title: row[0] || '',
    summary: row[1] || '',
    deadline: row[2] ? formatDate(row[2]) : '',
    mailLink: row[3] || '',
    category: row[4] || ''
  }));
}

// 総イベント数を取得
function getTotalEventCount(sheetName) {
  const sheetId = getSheetId();
  if (!sheetId) {
    return 0;
  }
  
  const ss = SpreadsheetApp.openById(sheetId);
  const sheet = ss.getSheetByName(sheetName);
  return sheet ? sheet.getLastRow() : 0;
}

// 日付フォーマット
function formatDate(date) {
  if (!date) return '';
  
  if (date instanceof Date) {
    const y = date.getFullYear();
    const m = ('0' + (date.getMonth() + 1)).slice(-2);
    const d = ('0' + date.getDate()).slice(-2);
    return `${y}-${m}-${d}`;
  }
  
  return date.toString();
}

// 成功レスポンス作成
function createSuccessResponse(data) {
  const output = ContentService
    .createTextOutput(JSON.stringify({
      success: true,
      data: data
    }))
    .setMimeType(ContentService.MimeType.JSON);
  
  return output;
}

// エラーレスポンス作成
function createErrorResponse(message, code) {
  const output = ContentService
    .createTextOutput(JSON.stringify({
      success: false,
      error: message,
      code: code
    }))
    .setMimeType(ContentService.MimeType.JSON);
  
  return output;
}

// POSTリクエスト処理（将来の拡張用）
function doPost(e) {
  try {
    // APIキー認証
    const apiKey = getApiKey();
    const requestApiKey = e.parameter.apiKey || JSON.parse(e.postData.contents).apiKey;
    if (apiKey && requestApiKey !== apiKey) {
      return createErrorResponse('Unauthorized', 401);
    }
    
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    
    switch (action) {
      case 'addEvent':
        // 将来的にイベント追加機能を実装
        return createErrorResponse('Not implemented', 501);
      default:
        return createErrorResponse('Invalid action', 400);
    }
    
  } catch (error) {
    console.error('Error:', error);
    return createErrorResponse(error.toString(), 500);
  }
}

// テスト用関数
function testAPI() {
  // テスト用のモックリクエスト
  const mockRequest = {
    parameter: {
      action: 'list',
      limit: '10'
    }
  };
  
  const response = doGet(mockRequest);
  console.log(response.getContent());
}