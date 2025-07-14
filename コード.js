// Gmail to Notion Integration Script

// グローバル変数
const SETTINGS_SHEET_NAME = '設定シート';
const NOTION_API_VERSION = '2022-06-28';

/**
 * スプレッドシート開始時に実行される関数
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  // カスタムメニューを作成
  ui.createMenu('Gmail-Notion連携')
    .addItem('設定シートを作成', 'createSettingsSheet')
    .addItem('手動で同期を実行', 'manualSync')
    .addItem('1時間ごとのトリガーを設定', 'setHourlyTrigger')
    .addItem('トリガーを削除', 'deleteTriggers')
    .addSeparator()
    .addItem('設定を入力', 'showSettingsDialog')
    .addToUi();
}

/**
 * 設定シートを作成する関数
 */
function createSettingsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 既存の設定シートがあるか確認
  let settingsSheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
  
  if (settingsSheet) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      '確認',
      '設定シートは既に存在します。新しく作り直しますか？',
      ui.ButtonSet.YES_NO
    );
    
    if (response === ui.Button.NO) {
      return;
    }
    
    // 既存のシートを削除
    ss.deleteSheet(settingsSheet);
  }
  
  // 新しい設定シートを作成
  settingsSheet = ss.insertSheet(SETTINGS_SHEET_NAME);
  
  // ヘッダーを設定
  const headers = [
    ['設定項目', '値', '説明'],
    ['Notion インテグレーションキー', '', 'NotionのInternal Integration Tokenを入力'],
    ['Notion データベースID', '', 'データベースのIDを入力（URLから取得）'],
    ['タイトルプロパティ名', 'タイトル', 'Notionデータベースのタイトル列の名前'],
    ['Gmailラベル名', 'Notion', '処理対象のGmailラベル名'],
    ['最終処理日時', '', '自動更新されます（変更不要）']
  ];
  
  // データを設定
  settingsSheet.getRange(1, 1, headers.length, 3).setValues(headers);
  
  // フォーマット設定
  settingsSheet.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#e0e0e0');
  settingsSheet.setColumnWidth(1, 200);
  settingsSheet.setColumnWidth(2, 300);
  settingsSheet.setColumnWidth(3, 400);
  
  // 保護設定（ヘッダー行）
  const protection = settingsSheet.getRange(1, 1, 1, 3).protect();
  protection.setDescription('ヘッダー行の保護');
  
  SpreadsheetApp.getUi().alert('設定シートを作成しました。設定項目の「値」列に必要な情報を入力してください。');
}

/**
 * 設定入力ダイアログを表示
 */
function showSettingsDialog() {
  const html = HtmlService.createHtmlOutputFromFile('settings-dialog')
    .setWidth(500)
    .setHeight(400);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Gmail-Notion連携設定');
}

/**
 * 設定を取得する関数
 */
function getSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
  
  if (!settingsSheet) {
    throw new Error('設定シートが見つかりません。先に設定シートを作成してください。');
  }
  
  const data = settingsSheet.getRange(2, 1, 5, 2).getValues();
  
  return {
    notionKey: data[0][1],
    databaseId: data[1][1],
    titleProperty: data[2][1] || 'タイトル',
    gmailLabel: data[3][1] || 'Notion',
    lastProcessed: data[4][1] ? new Date(data[4][1]) : null
  };
}

/**
 * 設定を保存する関数
 */
function saveSettings(settings) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
  
  if (!settingsSheet) {
    createSettingsSheet();
    return saveSettings(settings);
  }
  
  // 設定を書き込み
  if (settings.notionKey !== undefined) {
    settingsSheet.getRange(2, 2).setValue(settings.notionKey);
  }
  if (settings.databaseId !== undefined) {
    settingsSheet.getRange(3, 2).setValue(settings.databaseId);
  }
  if (settings.titleProperty !== undefined) {
    settingsSheet.getRange(4, 2).setValue(settings.titleProperty);
  }
  if (settings.gmailLabel !== undefined) {
    settingsSheet.getRange(5, 2).setValue(settings.gmailLabel);
  }
  if (settings.lastProcessed !== undefined) {
    settingsSheet.getRange(6, 2).setValue(settings.lastProcessed);
  }
}

/**
 * メールをNotionに同期するメイン関数
 */
function syncEmailsToNotion() {
  try {
    const settings = getSettings();
    
    // 設定の検証
    if (!settings.notionKey || !settings.databaseId) {
      throw new Error('NotionのインテグレーションキーとデータベースIDを設定してください。');
    }
    
    // Gmailラベルからメールを取得
    const label = GmailApp.getUserLabelByName(settings.gmailLabel);
    if (!label) {
      throw new Error(`ラベル「${settings.gmailLabel}」が見つかりません。`);
    }
    
    // 最終処理日時以降のメールを取得
    let query = `label:${settings.gmailLabel}`;
    if (settings.lastProcessed) {
      const dateStr = Utilities.formatDate(settings.lastProcessed, 'GMT', 'yyyy/MM/dd');
      query += ` after:${dateStr}`;
    }
    
    const threads = GmailApp.search(query, 0, 50); // 最大50件
    
    if (threads.length === 0) {
      console.log('新しいメールはありません。');
      return;
    }
    
    let processedCount = 0;
    let lastProcessedTime = settings.lastProcessed || new Date(0);
    
    // 各スレッドを処理
    for (const thread of threads) {
      const messages = thread.getMessages();
      
      for (const message of messages) {
        const messageDate = message.getDate();
        
        // 既に処理済みの日時より古い場合はスキップ
        if (settings.lastProcessed && messageDate <= settings.lastProcessed) {
          continue;
        }
        
        // Notionにページを作成
        const success = createNotionPage(settings, message);
        
        if (success) {
          processedCount++;
          if (messageDate > lastProcessedTime) {
            lastProcessedTime = messageDate;
          }
        }
      }
    }
    
    // 最終処理日時を更新
    saveSettings({ lastProcessed: lastProcessedTime });
    
    console.log(`${processedCount}件のメールをNotionに追加しました。`);
    
    return processedCount;
    
  } catch (error) {
    console.error('エラーが発生しました:', error);
    throw error;
  }
}

/**
 * Notionにページを作成する関数
 */
function createNotionPage(settings, message) {
  const url = `https://api.notion.com/v1/pages`;
  
  // メール本文を取得（プレーンテキスト優先）
  let body = message.getPlainBody();
  if (!body) {
    body = message.getBody();
    // HTMLタグを除去
    body = body.replace(/<[^>]*>/g, '');
  }
  
  // 本文を2000文字に制限（Notion APIの制限）
  if (body.length > 2000) {
    body = body.substring(0, 1997) + '...';
  }
  
  // リクエストボディを作成
  const requestBody = {
    parent: {
      database_id: settings.databaseId
    },
    properties: {
      [settings.titleProperty]: {
        title: [
          {
            text: {
              content: message.getSubject() || '（件名なし）'
            }
          }
        ]
      }
    },
    children: [
      {
        object: 'block',
        type: 'heading_2',
        heading_2: {
          rich_text: [
            {
              type: 'text',
              text: {
                content: 'メール情報'
              }
            }
          ]
        }
      },
      {
        object: 'block',
        type: 'paragraph',
        paragraph: {
          rich_text: [
            {
              type: 'text',
              text: {
                content: `送信者: ${message.getFrom()}\n日時: ${message.getDate()}`
              }
            }
          ]
        }
      },
      {
        object: 'block',
        type: 'divider',
        divider: {}
      },
      {
        object: 'block',
        type: 'heading_2',
        heading_2: {
          rich_text: [
            {
              type: 'text',
              text: {
                content: '本文'
              }
            }
          ]
        }
      },
      {
        object: 'block',
        type: 'paragraph',
        paragraph: {
          rich_text: [
            {
              type: 'text',
              text: {
                content: body
              }
            }
          ]
        }
      }
    ]
  };
  
  // Notion APIにリクエスト
  const options = {
    method: 'post',
    headers: {
      'Authorization': `Bearer ${settings.notionKey}`,
      'Content-Type': 'application/json',
      'Notion-Version': NOTION_API_VERSION
    },
    payload: JSON.stringify(requestBody),
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const result = JSON.parse(response.getContentText());
    
    if (response.getResponseCode() !== 200) {
      console.error('Notion APIエラー:', result);
      return false;
    }
    
    return true;
    
  } catch (error) {
    console.error('リクエストエラー:', error);
    return false;
  }
}

/**
 * 手動同期を実行
 */
function manualSync() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const count = syncEmailsToNotion();
    ui.alert('同期完了', `${count}件のメールをNotionに追加しました。`, ui.ButtonSet.OK);
  } catch (error) {
    ui.alert('エラー', `同期中にエラーが発生しました: ${error.message}`, ui.ButtonSet.OK);
  }
}

/**
 * 1時間ごとのトリガーを設定
 */
function setHourlyTrigger() {
  // 既存のトリガーを削除
  deleteTriggers();
  
  // 新しいトリガーを作成
  ScriptApp.newTrigger('syncEmailsToNotion')
    .timeBased()
    .everyHours(1)
    .create();
  
  SpreadsheetApp.getUi().alert('1時間ごとの自動同期を設定しました。');
}

/**
 * すべてのトリガーを削除
 */
function deleteTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'syncEmailsToNotion') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

// ===== 以下、HTMLダイアログ用の関数 =====

/**
 * 現在の設定を取得（HTMLダイアログ用）
 */
function getCurrentSettings() {
  try {
    return getSettings();
  } catch (error) {
    return {
      notionKey: '',
      databaseId: '',
      titleProperty: 'タイトル',
      gmailLabel: 'Notion'
    };
  }
}

/**
 * 設定を更新（HTMLダイアログ用）
 */
function updateSettings(newSettings) {
  saveSettings(newSettings);
  return true;
}