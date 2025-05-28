// デバッガーを使った簡単なテスト
function testWithDebugger() {
  console.log('デバッグ開始');
  
  debugger; // ここで一時停止します
  
  // 現在の設定を確認
  const sheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  console.log('SPREADSHEET_ID:', sheetId);
  
  debugger; // ここでも一時停止
  
  if (sheetId) {
    try {
      const ss = SpreadsheetApp.openById(sheetId);
      console.log('スプレッドシート名:', ss.getName());
      
      debugger; // 成功した場合ここで停止
      
      // シート一覧を取得
      const sheets = ss.getSheets();
      console.log('シート数:', sheets.length);
      
    } catch (error) {
      debugger; // エラーの場合ここで停止
      console.log('エラー:', error.toString());
    }
  } else {
    console.log('SPREADSHEET_IDが設定されていません');
  }
  
  console.log('デバッグ終了');
}

// まず実行する関数（通常実行）
function runFirst() {
  console.log('===== セットアップ開始 =====');
  
  // 1. 現在の設定確認
  const props = PropertiesService.getScriptProperties().getProperties();
  console.log('現在の設定:', JSON.stringify(props, null, 2));
  
  if (!props.SPREADSHEET_ID) {
    console.log('\n❌ SPREADSHEET_IDが設定されていません');
    console.log('次の手順を実行してください:');
    console.log('1. findMySpreadsheet() を実行');
    console.log('2. setupSpreadsheetId("ここにID") を実行');
  } else {
    console.log('\n✅ SPREADSHEET_ID:', props.SPREADSHEET_ID);
    testConnection();
  }
}

// スプレッドシートを探す
function findMySpreadsheet() {
  console.log('===== スプレッドシートを検索 =====');
  
  try {
    const files = DriveApp.searchFiles('mimeType = "application/vnd.google-apps.spreadsheet"');
    let count = 0;
    
    console.log('\n最近のスプレッドシート:');
    while (files.hasNext() && count < 15) {
      const file = files.next();
      const name = file.getName();
      const id = file.getId();
      
      // イベント関連のキーワード
      if (name.match(/イベント|event|Event|N高|締切|課外活動/i)) {
        console.log('\n⭐ 【おそらくこれ】');
        console.log('  名前: ' + name);
        console.log('  ID: ' + id);
        console.log('  → setupSpreadsheetId("' + id + '") を実行してください');
      } else {
        console.log('・ ' + name + ' (ID: ' + id + ')');
      }
      count++;
    }
  } catch (error) {
    console.log('エラー:', error.toString());
  }
}

// スプレッドシートIDを設定
function setupSpreadsheetId(id) {
  if (!id || id === 'ここにID') {
    console.log('使い方: setupSpreadsheetId("実際のスプレッドシートID")');
    return;
  }
  
  console.log('===== スプレッドシートID設定 =====');
  
  try {
    // アクセステスト
    const ss = SpreadsheetApp.openById(id);
    console.log('✅ アクセス成功: ' + ss.getName());
    
    // 保存
    PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', id);
    console.log('✅ 保存完了');
    
    // 接続テスト
    testConnection();
    
  } catch (error) {
    console.log('❌ エラー: ' + error.toString());
  }
}

// 接続テスト
function testConnection() {
  console.log('\n===== 接続テスト =====');
  
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) {
    console.log('❌ SPREADSHEET_IDが設定されていません');
    return;
  }
  
  try {
    const ss = SpreadsheetApp.openById(id);
    console.log('スプレッドシート: ' + ss.getName());
    console.log('\nシート一覧:');
    
    ss.getSheets().forEach(sheet => {
      const name = sheet.getName();
      const rows = sheet.getLastRow();
      console.log('  ・' + name + ' (' + rows + '行のデータ)');
      
      // イベント一覧シートがあれば最初の1行を表示
      if (name === 'イベント一覧' && rows > 0) {
        const firstRow = sheet.getRange(1, 1, 1, 5).getValues()[0];
        console.log('    最初のイベント: ' + firstRow[0]);
      }
    });
    
    console.log('\n✅ すべて正常です！');
    console.log('→ WebアプリとしてデプロイしてAPIを使用できます');
    
  } catch (error) {
    console.log('❌ エラー: ' + error.toString());
  }
}