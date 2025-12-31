/**
 * MODA AI 行政秘書 - Google Apps Script API
 * 處理網頁與 Google Sheets 的資料寫入
 */

// 允許跨網域存取
function doPost(e) {
  return handleRequest(e);
}

function doGet(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  // 設定 CORS 標頭
  const output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);
  
  try {
    const params = e.parameter;
    const action = params.action;
    
    let result;
    
    switch(action) {
      case 'addMeeting':
        result = addMeeting(params);
        break;
      case 'addTodo':
        result = addTodo(params);
        break;
      case 'updateStatus':
        result = updateStatus(params);
        break;
      case 'getNextId':
        result = getNextId(params.sheet);
        break;
      default:
        result = { success: false, error: '未知的操作' };
    }
    
    output.setContent(JSON.stringify(result));
    
  } catch (error) {
    output.setContent(JSON.stringify({ 
      success: false, 
      error: error.toString() 
    }));
  }
  
  return output;
}

/**
 * 新增會議/工作
 */
function addMeeting(params) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('01_會議工作清單');
  
  if (!sheet) {
    return { success: false, error: '找不到會議工作清單工作表' };
  }
  
  // 產生新編號
  const lastRow = sheet.getLastRow();
  const year = new Date().getFullYear();
  const newId = `MTG-${year}-${String(lastRow).padStart(3, '0')}`;
  
  // 新增資料列
  const newRow = [
    newId,                          // 編號
    params.title || '',             // 主題
    params.category || '',          // 工作分類
    params.organization || '',      // 相關單位
    params.assignee || '',          // 負責人
    params.assignDate || '',        // 指派日期
    params.dueDate || '',           // 截止日期
    '',                             // 完成日期
    '',                             // 處理天數（公式會自動計算）
    params.fileCount || 0,          // 提交檔案數
    params.status || '待處理',      // 狀態
    params.note || ''               // 備註
  ];
  
  sheet.appendRow(newRow);
  
  return { 
    success: true, 
    message: '會議已新增',
    meetingId: newId 
  };
}

/**
 * 新增待辦事項
 */
function addTodo(params) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('05_待辦追蹤');
  
  if (!sheet) {
    return { success: false, error: '找不到待辦追蹤工作表' };
  }
  
  const lastRow = sheet.getLastRow();
  const newId = `TODO-${String(lastRow).padStart(3, '0')}`;
  
  const newRow = [
    newId,                          // 待辦編號
    params.meetingId || '',         // 關聯會議編號
    params.task || '',              // 待辦事項
    params.assignee || '',          // 負責人
    params.assigner || '',          // 指派人
    params.createDate || new Date().toISOString().split('T')[0],
    params.dueDate || '',           // 截止日期
    '',                             // 完成日期
    params.status || '待處理',      // 狀態
    params.priority || '中'         // 優先級
  ];
  
  sheet.appendRow(newRow);
  
  return { 
    success: true, 
    message: '待辦已新增',
    todoId: newId 
  };
}

/**
 * 更新狀態
 */
function updateStatus(params) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheetName = params.sheet || '01_會議工作清單';
  
  // 相容舊名稱
  if (sheetName === '會議工作清單') sheetName = '01_會議工作清單';
  if (sheetName === '待辦追蹤') sheetName = '05_待辦追蹤';
  
  const sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    return { success: false, error: '找不到工作表' };
  }
  
  const id = params.id;
  const newStatus = params.status;
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      // 找到該列，更新狀態欄（第11欄，索引10）
      const statusCol = sheetName === '05_待辦追蹤' ? 9 : 11;
      sheet.getRange(i + 1, statusCol).setValue(newStatus);
      
      // 如果是完成狀態，記錄完成日期
      if (newStatus === '已完成') {
        const completeDateCol = sheetName === '05_待辦追蹤' ? 8 : 8;
        sheet.getRange(i + 1, completeDateCol).setValue(new Date().toISOString().split('T')[0]);
      }
      
      return { success: true, message: '狀態已更新' };
    }
  }
  
  return { success: false, error: '找不到指定的項目' };
}

/**
 * 取得下一個編號
 */
function getNextId(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 相容舊名稱
  if (!sheetName || sheetName === '會議工作清單') sheetName = '01_會議工作清單';
  if (sheetName === '待辦追蹤') sheetName = '05_待辦追蹤';
  
  const sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    return { success: false, error: '找不到工作表' };
  }
  
  const lastRow = sheet.getLastRow();
  const year = new Date().getFullYear();
  const prefix = sheetName === '05_待辦追蹤' ? 'TODO' : `MTG-${year}`;
  
  return { 
    success: true, 
    nextId: `${prefix}-${String(lastRow).padStart(3, '0')}` 
  };
}
