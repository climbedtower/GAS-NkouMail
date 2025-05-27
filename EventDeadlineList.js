/******************** 設定 ********************/
const OPENAI_API_KEY = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
const SHEET_ID       = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID'); // 出力先シート
const LOOKBACK_DAYS  = Number(PropertiesService.getScriptProperties().getProperty('LOOKBACK_DAYS')) || 1; // 何日分のメールを取得するか
const MODEL_CHEAP    = 'gpt-3.5-turbo-0125';  // 構造抽出用
const MODEL_SUMMARY  = 'gpt-4o-mini';         // 要約用
// イベントカテゴリ判定用モデル
const MODEL_CATEGORY = 'gpt-4o-mini';         // カテゴリ判定用
const MAILS_PER_BATCH = 5;                    // まとめて投げる通数
const SHEET_EVENTS   = 'イベント一覧';        // イベント一覧シート名
const CATEGORIES     = ['課外授業', '重要/テスト', 'その他'];

// 件名や本文に基づくキーワード判定
function guessCategory(text) {
  if (!text) return 'その他';
  const t = text.toLowerCase();
  const important = /(締切|提出|試験|テスト|成績|重要|レポート)/i;
  const extracurricular = /(課外|体験学習|ワークショップ|交流|イベント|ゲーム)/i;
  if (important.test(t)) return '重要/テスト';
  if (extracurricular.test(t)) return '課外授業';
  return 'その他';
}

/****************** シート既存取得 & 書き込み ******************/
function getExistingKeys(sheetName) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName(sheetName);
  // データ行が存在しない場合は空セットを返す
  if (!sh || sh.getLastRow() < 1) return new Set();
  const numRows = sh.getLastRow();
  const data = sh.getRange(1, 1, numRows, 5).getValues();
  // 既存データは件名(1列目)と締切(3列目)のみでキー生成
  return new Set(
    data.map(r => [r[0], normalizeDeadlineFormat(r[2])].join('|'))
  );
}

function writeRowsUnique(sheetName, rows, existingKeys, splitByCategory) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName(sheetName);
  if (!sh) sh = ss.insertSheet(sheetName);
  const newRows = rows.filter(r => {
    const key = [r[0], r[2]].join('|'); // 件名と締切のみで判断
    if (existingKeys.has(key)) return false;
    existingKeys.add(key);
    return true;
  });
  if (newRows.length) {
    const startRow = sh.getLastRow() + 1;
    sh.getRange(startRow, 1, newRows.length, 5).setValues(newRows);
    Logger.log(`シート「${sheetName}」に ${newRows.length} 行追加`);
  } else {
    Logger.log(`シート「${sheetName}」に追加行なし`);
  }

  if (splitByCategory && newRows.length) {
    const grouped = {};
    newRows.forEach(r => {
      const cat = r[4] || '未分類';
      const catSheet = cat;
      if (!grouped[catSheet]) grouped[catSheet] = [];
      grouped[catSheet].push(r);
    });
    for (const [catSheet, rows] of Object.entries(grouped)) {
      let csh = ss.getSheetByName(catSheet);
      if (!csh) csh = ss.insertSheet(catSheet);
      const start = csh.getLastRow() + 1;
      csh.getRange(start, 1, rows.length, 5).setValues(rows);
      Logger.log(`シート「${catSheet}」に ${rows.length} 行追加`);
    }
  }
}

/******************** 重複排除 (期限＋内容) ********************/
function levenshtein(a, b) {
  const m = a.length, n = b.length;
  const d = Array.from({ length: m + 1 }, () => Array(n + 1).fill(0));
  for (let i = 0; i <= m; i++) d[i][0] = i;
  for (let j = 0; j <= n; j++) d[0][j] = j;
  for (let i = 1; i <= m; i++) for (let j = 1; j <= n; j++) {
    const cost = a[i - 1] === b[j - 1] ? 0 : 1;
    d[i][j] = Math.min(d[i - 1][j] + 1, d[i][j - 1] + 1, d[i - 1][j - 1] + cost);
  }
  return d[m][n];
}
function similarity(a, b) {
  if (!a || !b) return 0;
  const dist = levenshtein(a, b);
  return 1 - dist / Math.max(a.length, b.length);
}
function dedupeEvents(events) {
  const seen = {};
  const result = [];

  events.forEach(ev => {
    const dateKey = ev.deadline || '_none_';
    seen[dateKey] = seen[dateKey] || [];

    let merged = false;
    for (const existing of seen[dateKey]) {
      if (similarity(existing.title, ev.title) > 0.8) {
        if (!existing.summary && ev.summary) {
          existing.summary = ev.summary;
          existing.row[1] = ev.summary;
        }
        if (ev.mailDate && (!existing.mailDate || ev.mailDate > existing.mailDate)) {
          existing.mailDate = ev.mailDate;
          existing.row[3] = ev.row[3];
        }
        if (!existing.category && ev.category) {
          existing.category = ev.category;
          existing.row[4] = ev.category;
        }
        if (!existing.subject && ev.subject) {
          existing.subject = ev.subject;
        }
        if (!existing.body && ev.body) {
          existing.body = ev.body;
        }
        if (!existing.preCategory && ev.preCategory) {
          existing.preCategory = ev.preCategory;
        }
        merged = true;
        break;
      }
    }

    if (!merged) {
      seen[dateKey].push(ev);
      result.push(ev);
    }
  });

  return result;
}

// キーワードからの暫定カテゴリ付与
function applyKeywordCategory(events) {
  events.forEach(ev => {
    if (!ev.preCategory) {
      const text = [ev.title, ev.subject, ev.body].filter(Boolean).join('\n');
      ev.preCategory = guessCategory(text);
    }
  });
}

/******************** メイン ********************/
function summarizeNHighEmails() {
  try {
    const existing = getExistingKeys(SHEET_EVENTS);

    const threads = GmailApp.search(`N高 newer_than:${LOOKBACK_DAYS}d`, 0, 200);
    const messages = threads.reduce((all, th) => all.concat(th.getMessages()), []);
    Logger.log(`対象メール ${messages.length} 通`);
    if (!messages.length) return;

    let events = [];
    for (let i = 0; i < messages.length; i += MAILS_PER_BATCH) {
      events.push(...extractBatchEvents(messages.slice(i, i + MAILS_PER_BATCH)));
      if (i + MAILS_PER_BATCH < messages.length) Utilities.sleep(1000);
    }
    Logger.log(`抽出イベント(前重複排除): ${events.length} 件`);

    events = dedupeEvents(events);
    applyKeywordCategory(events);
    Logger.log(`抽出イベント(重複排除後): ${events.length} 件`);

  const needSummary = events.filter(e => e.title && !e.summary);
  Logger.log(`要約対象イベント: ${needSummary.length} 件`);
  for (let i = 0; i < needSummary.length; i += MAILS_PER_BATCH) {
    summarizeBatch(needSummary.slice(i, i + MAILS_PER_BATCH));
    if (i + MAILS_PER_BATCH < needSummary.length) Utilities.sleep(1000);
  }

  const needCategory = events.filter(e => !e.category);
  Logger.log(`カテゴリ付与対象イベント: ${needCategory.length} 件`);
  for (let i = 0; i < needCategory.length; i += MAILS_PER_BATCH) {
    categorizeBatch(needCategory.slice(i, i + MAILS_PER_BATCH));
    if (i + MAILS_PER_BATCH < needCategory.length) Utilities.sleep(1000);
  }

    events.sort((a, b) => {
      if (a.deadline && b.deadline) {
        return new Date(a.deadline) - new Date(b.deadline);
      }
      if (a.deadline) return -1;
      if (b.deadline) return 1;
      return 0;
    });

    writeRowsUnique(SHEET_EVENTS, events.map(e => e.row), existing, true);

    Logger.log('処理完了');
  } catch (e) {
    Logger.log('エラー: ' + e);
    throw e;
  }
} 

/**************** 日付抽出ユーティリティ *****************/
function extractDeadlineFromText(text) {
  if (!text) return '';

  let m = text.match(/(\d{4})[\/\.](\d{1,2})[\/\.](\d{1,2})/);
  if (m) {
    const y = m[1];
    const mo = ('0' + m[2]).slice(-2);
    const d = ('0' + m[3]).slice(-2);
    return `${y}-${mo}-${d}`;
  }

  m = text.match(/(\d{4})年\s*(\d{1,2})月\s*(\d{1,2})日/);
  if (m) {
    const y = m[1];
    const mo = ('0' + m[2]).slice(-2);
    const d = ('0' + m[3]).slice(-2);
    return `${y}-${mo}-${d}`;
  }

  m = text.match(/(\d{1,2})月\s*(\d{1,2})日/);
  if (m) {
    const y = new Date().getFullYear();
    const mo = ('0' + m[1]).slice(-2);
    const d = ('0' + m[2]).slice(-2);
    return `${y}-${mo}-${d}`;
  }

  m = text.match(/(\d{1,2})[\/\.](\d{1,2})/);
  if (m) {
    const y = new Date().getFullYear();
    const mo = ('0' + m[1]).slice(-2);
    const d = ('0' + m[2]).slice(-2);
    return `${y}-${mo}-${d}`;
  }

  return '';
}

function normalizeDeadlineFormat(dateStr) {
  if (dateStr instanceof Date) {
    const y = dateStr.getFullYear();
    const mo = ('0' + (dateStr.getMonth() + 1)).slice(-2);
    const day = ('0' + dateStr.getDate()).slice(-2);
    return `${y}-${mo}-${day}`;
  }

  if (!dateStr) return '';

  if (typeof dateStr === 'string' && dateStr.trim().toUpperCase() === 'YYYY-MM-DD') {
    return '';
  }

  const detected = extractDeadlineFromText(dateStr);
  if (detected) return detected;

  const d = new Date(dateStr);
  if (!isNaN(d)) {
    const y = d.getFullYear();
    const mo = ('0' + (d.getMonth() + 1)).slice(-2);
    const day = ('0' + d.getDate()).slice(-2);
    return `${y}-${mo}-${day}`;
  }

  return '';
}

/**************** AI でイベント抽出 ***************/
/* ── 3.1 タイトル・締切抽出 (安価モデル) ───────────────── */
function extractBatchEvents(msgArray) {
  try {
    const promptParts = msgArray.map((m, idx) => {
      const subject = m.getSubject() || 'タイトルなし';
      const body = m.getPlainBody().slice(0, 3000); // 少し短くしてトークン制限回避
      return `### メール${idx + 1}\n件名: ${subject}\n本文:\n${body}`;
    });
    
    const userPrompt = `
次の複数メールから「イベント名・締切日」を抽出し、メール毎に JSON を返してください。
出力形式:
[
  {"mailIndex":1,"events":[
   {"title":"イベント名","deadline":"2025-12-31 または空文字","hasDeadline":true}
 ]}
]

重要な注意点:
- 締切が明確でない場合は deadline を空文字 "" にする
- 締切日は YYYY-MM-DD 形式で統一する
- イベントが見つからない場合は events を空の配列 [] にする
- 必ず有効なJSONフォーマットで返す

-----
${promptParts.join('\n\n')}
-----`;

    Logger.log('OpenAI API 呼び出し開始（抽出）');
    const resTxt = openaiCall(MODEL_CHEAP, userPrompt);
    Logger.log('▽ 抽出結果 RAW（最初の500字）:\n' + resTxt.slice(0, 500));
    
    const arr = safeJsonParse(resTxt);
    if (!Array.isArray(arr)) {
      Logger.log('警告: JSON解析結果が配列ではありません');
      return [];
    }

    /* arr をイベント単位に展開 */
    const out = [];
    arr.forEach((obj, objIndex) => {
      if (!obj.mailIndex || obj.mailIndex < 1 || obj.mailIndex > msgArray.length) {
        Logger.log(`警告: 無効なmailIndex: ${obj.mailIndex}`);
        return;
      }
      
      const mail = msgArray[obj.mailIndex - 1];
      const link = mailLink(mail);
      const defaultTitle = mail.getSubject() || 'タイトルなし';
      
      if (!obj.events || !Array.isArray(obj.events)) {
        Logger.log(`警告: メール${obj.mailIndex}のeventsが配列ではありません`);
        return;
      }
      
        obj.events.forEach((ev, evIndex) => {
          const title = ev.title || defaultTitle;
          let deadline = normalizeDeadlineFormat(ev.deadline || '');

          if (!deadline) {
            const detected = extractDeadlineFromText(mail.getSubject() + '\n' + mail.getPlainBody());
            if (detected) deadline = detected;
          }

          const subject = mail.getSubject() || '';
          const body = mail.getPlainBody();
          out.push({
            title: title,
            summary: '',              // 後で埋める
            deadline: deadline,
            category: '',            // 後で埋める
            preCategory: guessCategory(subject + '\n' + body),
            mailDate: mail.getDate(),
            subject: subject,
            body: body.slice(0, 1000),
            row: [title, '', deadline, link, '']
          });
        });
    });
    
    Logger.log(`バッチから ${out.length} 件のイベントを抽出`);
    return out;
    
  } catch (error) {
    Logger.log('extractBatchEvents でエラー: ' + error.toString());
    return [];
  }
}

/* ── 3.2 要約 (GPT-4o mini) ───────────────────────────── */
function summarizeBatch(eventArr) {
  try {
    if (!eventArr || eventArr.length === 0) {
      Logger.log('要約対象のイベントがありません');
      return;
    }
    
    const combined = eventArr.map((ev, idx) => {
      const title = ev.title || 'タイトルなし';
      return `【${idx + 1}】${title}`;
    }).join('\n');
    
    const prompt = `以下のイベントタイトルを各90字以内で日本語要約し JSON 配列で返してください:

出力形式:
[
 {"index":1,"summary":"要約文"}
]

イベント一覧:
${combined}`;

    Logger.log('OpenAI API 呼び出し開始（要約）');
    const resTxt = openaiCall(MODEL_SUMMARY, prompt);
    const arr = safeJsonParse(resTxt);
    
    if (!Array.isArray(arr)) {
      Logger.log('警告: 要約結果のJSON解析に失敗');
      return;
    }
    
    arr.forEach(o => {
      if (!o.index || o.index < 1 || o.index > eventArr.length) {
        Logger.log(`警告: 無効な要約index: ${o.index}`);
        return;
      }
      
      const ev = eventArr[o.index - 1];
      ev.summary = o.summary || '';
      ev.row[1] = o.summary || '';
    });
    
    Logger.log(`${arr.length} 件の要約を処理`);
    
  } catch (error) {
    Logger.log('summarizeBatch でエラー: ' + error.toString());
  }
}

function categorizeBatch(eventArr) {
  try {
    if (!eventArr || eventArr.length === 0) {
      Logger.log('カテゴリ付与対象のイベントがありません');
      return;
    }

    const combined = eventArr.map((ev, idx) => {
      const title = ev.title || 'タイトルなし';
      return `【${idx + 1}】${title}`;
    }).join('\n');

    const prompt = `以下のイベントタイトルを次のいずれかのカテゴリに分類し JSON 配列で返してください: ${CATEGORIES.join(', ')}\n\n出力形式:\n[\n {"index":1,"category":"課外授業"}\n]\n\nイベント一覧:\n${combined}`;

    Logger.log('OpenAI API 呼び出し開始（カテゴリ）');
    const resTxt = openaiCall(MODEL_CATEGORY, prompt);
    const arr = safeJsonParse(resTxt);
    if (!Array.isArray(arr)) {
      Logger.log('警告: カテゴリ結果のJSON解析に失敗');
      return;
    }

    arr.forEach(o => {
      if (!o.index || o.index < 1 || o.index > eventArr.length) {
        Logger.log(`警告: 無効なカテゴリindex: ${o.index}`);
        return;
      }

      const ev = eventArr[o.index - 1];
      const aiCat = CATEGORIES.includes(o.category) ? o.category : '';
      let finalCat = aiCat;
      if (!finalCat || finalCat === 'その他') {
        if (ev.preCategory && ev.preCategory !== 'その他') {
          finalCat = ev.preCategory;
        }
      }
      ev.category = finalCat || 'その他';
      ev.row[4] = ev.category;
    });

    // AI から返らなかったイベントに対しても補完
    eventArr.forEach(ev => {
      if (!ev.category) {
        ev.category = ev.preCategory || 'その他';
        ev.row[4] = ev.category;
      } else if (ev.category === 'その他' && ev.preCategory && ev.preCategory !== 'その他') {
        ev.category = ev.preCategory;
        ev.row[4] = ev.category;
      }
    });

    Logger.log(`${arr.length} 件のカテゴリを処理`);

  } catch (error) {
    Logger.log('categorizeBatch でエラー: ' + error.toString());
  }
}

function openaiCall(model, userPrompt) {
  const payload = {
    model: model,
    temperature: 0.2,
    messages: [
      { role: 'user', content: userPrompt }
    ]
  };
  
  for (let i = 0; i < 6; i++) {                       // 最大 6 回 (0→1→2→4→8→16s)
    try {
      const res = UrlFetchApp.fetch(
        'https://api.openai.com/v1/chat/completions',
        {
          method: 'post',
          headers: {
            'Authorization': 'Bearer ' + OPENAI_API_KEY,
            'Content-Type': 'application/json'
          },
          payload: JSON.stringify(payload),
          muteHttpExceptions: true
        }
      );
      
      const code = res.getResponseCode();
      const responseText = res.getContentText();
      
      Logger.log(`OpenAI API レスポンス: ${code}`);
      
      if (code === 200) {
        const jsonResponse = JSON.parse(responseText);
        if (jsonResponse.choices && jsonResponse.choices[0] && jsonResponse.choices[0].message) {
          return jsonResponse.choices[0].message.content.trim();
        } else {
          throw new Error('APIレスポンスの形式が不正です');
        }
      }
      
      if (code === 429) {
        const waitTime = 1000 * Math.pow(2, i);
        Logger.log(`レート制限により ${waitTime}ms 待機`);
        Utilities.sleep(waitTime);
        continue;
      }
      
      Logger.log('OpenAI APIエラー: ' + responseText);
      throw new Error(`OpenAI API ${code}: ${responseText}`);
      
    } catch (error) {
      if (i === 5) { // 最後のリトライ
        throw new Error('OpenAI API 呼び出しに失敗: ' + error.toString());
      }
      Logger.log(`リトライ ${i + 1}: ${error.toString()}`);
      Utilities.sleep(1000 * Math.pow(2, i));
    }
  }
  
  throw new Error('OpenAI retry exceeded');
}

function safeJsonParse(s) {
  try {
    if (!s) return [];

    // コードブロックが含まれる場合は除去
    s = s.trim();
    if (s.startsWith('```')) {
      s = s.replace(/^```\w*\n/, '').replace(/```\s*$/m, '').trim();
    }

    // まず全体をJSONとしてパース試行
    const parsed = JSON.parse(s);
    if (Array.isArray(parsed)) {
      return parsed;
    }
  } catch (_) {
    // パース失敗時は配列部分のみを抽出して再試行
    const arrText = extractJsonArray(s);
    if (arrText) {
      try {
        return JSON.parse(arrText);
      } catch (e) {
        Logger.log('JSON解析エラー: ' + e.toString());
        Logger.log('対象文字列: ' + arrText.slice(0, 200));
        return [];
      }
    }
    Logger.log('JSON解析失敗: 配列が見つかりません');
    Logger.log('対象文字列: ' + s.slice(0, 200));
    return [];
  }

  Logger.log('JSON解析失敗: 配列が見つかりません');
  Logger.log('対象文字列: ' + s.slice(0, 200));
  return [];
}

function extractJsonArray(str) {
  const start = str.indexOf('[');
  if (start === -1) return '';

  let depth = 0;
  for (let i = start; i < str.length; i++) {
    if (str[i] === '[') depth++;
    else if (str[i] === ']') {
      depth--;
      if (depth === 0) {
        return str.slice(start, i + 1);
      }
    }
  }
  return '';
}

function mailLink(m) {
  return `https://mail.google.com/mail/u/0/#inbox/${m.getId()}`;
}

function writeRows(name, rows) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sh = ss.getSheetByName(name);
    
    if (!sh) {
      sh = ss.insertSheet(name);
    }
    
    sh.clear();
    
    if (rows && rows.length > 0) {
      // 空の行をフィルタリング
      const validRows = rows.filter(row => row && row.length >= 5);
      if (validRows.length > 0) {
        sh.getRange(1, 1, validRows.length, 5).setValues(validRows);
      }
    }
    
    Logger.log(`シート「${name}」に ${rows ? rows.length : 0} 行を書き込み`);
    
  } catch (error) {
    Logger.log('シート書き込みエラー: ' + error.toString());
  }
}

// 期限切れイベントを削除する
function cleanupExpiredEvents() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheetNames = [SHEET_EVENTS, ...CATEGORIES];

  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const twoWeeksAgo = new Date(today);
  twoWeeksAgo.setDate(twoWeeksAgo.getDate() - 14);

  sheetNames.forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) return;

    const last = sh.getLastRow();
    if (last < 1) return;

    const values = sh.getRange(1, 1, last, 5).getValues();
    for (let i = last; i >= 1; i--) {
      const row = values[i - 1];
      const deadlineStr = normalizeDeadlineFormat(row[2]);
      const category = row[4];
      if (!deadlineStr) continue;
      const d = new Date(deadlineStr);
      d.setHours(0, 0, 0, 0);
      if (d < twoWeeksAgo && category !== CATEGORIES[1]) {
        sh.deleteRow(i);
      }
    }
  });
}
