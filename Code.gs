const TELEGRAM_BOT_TOKEN = 'xxxxxx:xxxxxxxxxxxxxxxxxxxxxxxxx';

// Function to serve the web app
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Deadline Monitoring Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Original checkDeadlines_3 function, updated for new column
function checkDeadlines_3() {
  Logger.log('Bắt đầu chạy checkDeadlines tại: ' + new Date().toLocaleString('vi-VN', {timeZone: 'Asia/Ho_Chi_Minh'}));
  
  var configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  if (!configSheet) {
    Logger.log('LỖI: Sheet "Config" không tồn tại.');
    return;
  }
  
  var data = configSheet.getDataRange().getValues();
  Logger.log('Đã đọc ' + (data.length - 1) + ' hàng từ sheet Config.');
  
  var today = new Date();
  today.setHours(0, 0, 0, 0); // Đặt về đầu ngày
  Logger.log('Ngày hiện tại: ' + Utilities.formatDate(today, 'GMT+7', 'yyyy-MM-dd'));
  
  for (var i = 1; i < data.length; i++) { // Bắt đầu từ hàng 2
    var row = data[i];
    var fileId = row[1].trim();
    var sheetName = row[3].trim() || null;
    var taskNameColLetter = row[4].trim().toUpperCase();
    var startColLetter = row[5].trim().toUpperCase();
    var endColLetter = row[6].trim().toUpperCase();
    var statusColLetter = row[7].trim().toUpperCase();
    var recipients = row[8].trim().split(',').map(e => e.trim());
    var sendMethod = row[9].trim().toLowerCase() || 'email'; // Mặc định Email
    var maxAlerts = parseInt(row[10]) || 0;
    var beforeDays = parseInt(row[11]) || 0;
    var afterDays = parseInt(row[12]) || 0;
    var conditions = row[13].trim() || '';
    var status = row[14].trim() || 'Đã gửi: 0 lần';
    var active = row[15].trim() || 'Yes';
    
    if (active !== 'Yes') {
      Logger.log('Bỏ qua hàng ' + (i + 1) + ': Không active.');
      continue;
    }
    
    Logger.log('Xử lý hàng ' + (i + 1) + ': File ID = ' + fileId + ', Sheet = ' + (sheetName || 'Default') + ', Recipients = ' + recipients.join(', ') + ', Phương thức = ' + sendMethod);
    
    if (!fileId) {
      Logger.log('Bỏ qua hàng ' + (i + 1) + ': Thiếu File ID.');
      continue;
    }
    
    // Kiểm tra phương thức gửi
    if (!['email', 'telegram', 'both'].includes(sendMethod)) {
      Logger.log('Bỏ qua hàng ' + (i + 1) + ': Phương thức gửi không hợp lệ (' + sendMethod + '). Mặc định dùng Email.');
      sendMethod = 'email';
    }
    
    // Kiểm tra danh sách email (cho Email hoặc Both)
    var validEmails = [];
    if (sendMethod === 'email' || sendMethod === 'both') {
      validEmails = recipients.filter(recipient => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(recipient));
      if (validEmails.length === 0) {
        Logger.log('Cảnh báo hàng ' + (i + 1) + ': Không có email hợp lệ (' + recipients.join(', ') + ').');
      }
    }
    
    // Kiểm tra danh sách Chat ID (cho Telegram hoặc Both)
    var validChatIds = [];
    if (sendMethod === 'telegram' || sendMethod === 'both') {
      validChatIds = recipients.filter(recipient => /^-?\d+$/.test(recipient));
      if (validChatIds.length === 0) {
        Logger.log('Cảnh báo hàng ' + (i + 1) + ': Không có Chat ID hợp lệ (' + recipients.join(', ') + ').');
      }
    }
    
    // Bỏ qua nếu không có email/Chat ID hợp lệ khi cần
    if ((sendMethod === 'email' && validEmails.length === 0) ||
        (sendMethod === 'telegram' && validChatIds.length === 0) ||
        (sendMethod === 'both' && validEmails.length === 0 && validChatIds.length === 0)) {
      Logger.log('Bỏ qua hàng ' + (i + 1) + ': Không có email hoặc Chat ID hợp lệ để gửi.');
      continue;
    }
    
    // Parse trạng thái hiện tại
    var sentCount = parseInt(status.match(/Đã gửi: (\d+) lần/)?.[1] || 0);
    if (sentCount >= maxAlerts && maxAlerts > 0) {
      Logger.log('Bỏ qua hàng ' + (i + 1) + ': Đã gửi đủ ' + maxAlerts + ' lần.');
      continue;
    }
    
    try {
      var targetSpreadsheet = SpreadsheetApp.openById(fileId);
      Logger.log('Mở file thành công: ' + targetSpreadsheet.getName() + ' (ID: ' + fileId + ')');
      
      var targetSheet = sheetName ? targetSpreadsheet.getSheetByName(sheetName) : targetSpreadsheet.getSheets()[0];
      if (!targetSheet) {
        Logger.log('LỖI: Không tìm thấy sheet ' + (sheetName || 'đầu tiên') + ' trong file ' + fileId);
        continue;
      }
      Logger.log('Đang xử lý sheet: ' + targetSheet.getName());
      
      var targetData = targetSheet.getDataRange().getValues();
      Logger.log('Đã đọc ' + targetData.length + ' hàng từ sheet ' + targetSheet.getName());
      
      var taskNameColIndex = columnLetterToIndex(taskNameColLetter);
      var startColIndex = columnLetterToIndex(startColLetter);
      var endColIndex = columnLetterToIndex(endColLetter);
      var statusColIndex = statusColLetter ? columnLetterToIndex(statusColLetter) : 0;
      
      // Parse điều kiện thêm
      var rowRange = parseRowRange(conditions);
      var skipConditions = parseConditions(conditions, 'Bỏ qua nếu');
      var checkConditions = parseConditions(conditions, 'Kiểm tra nếu');
      Logger.log('Điều kiện: Phạm vi hàng ' + rowRange[0] + ' đến ' + rowRange[1] + ', Bỏ qua = ' + JSON.stringify(skipConditions) + ', Kiểm tra = ' + JSON.stringify(checkConditions));
      
      var alerts = [];
      
      for (var j = (rowRange[0] - 1); j < Math.min(rowRange[1], targetData.length); j++) {
        var taskRow = targetData[j];
        var taskName = taskNameColIndex ? taskRow[taskNameColIndex - 1] || 'Task ở hàng ' + (j + 1) : 'Task ở hàng ' + (j + 1);
        var startDate = taskRow[startColIndex - 1];
        var endDate = taskRow[endColIndex - 1];
        var taskStatus = statusColIndex ? taskRow[statusColIndex - 1] : '';
        
        Logger.log('Kiểm tra hàng ' + (j + 1) + ': Task = ' + taskName + ', Status = ' + (taskStatus || 'N/A'));
        
        // Bỏ qua nếu startDate hoặc endDate rỗng
        if (!startDate || !endDate) {
          Logger.log('Bỏ qua hàng ' + (j + 1) + ': Start Date hoặc End Date rỗng (Start: ' + startDate + ', End: ' + endDate + ')');
          continue;
        }
        
        // Kiểm tra định dạng ngày
        function parseDate(str) {
        var parts = str.split('/');
        if (parts.length === 3) {
        // parts[0] = dd, parts[1] = mm, parts[2] = yy
        var day = parseInt(parts[0], 10);
        var month = parseInt(parts[1], 10) - 1; // tháng trong JS: 0-11
        var year = parseInt(parts[2], 10);
        if (year < 100) year += 2000; // convert 25 → 2025
        return new Date(year, month, day);
        }
        return null;
        }

        // Trong vòng lặp: Nếu đúng định dạng Date thì dùng luôn còn không thì convert sang DD/MM/YY
        if (!(startDate instanceof Date)) startDate = parseDate(startDate);
        if (!(endDate instanceof Date)) endDate = parseDate(endDate);
        
        if (!(startDate instanceof Date) || !(endDate instanceof Date)) {
          Logger.log('Bỏ qua hàng ' + (j + 1) + ': Ngày không hợp lệ (Start: ' + startDate + ', End: ' + endDate + ')');
          continue;
        }
        
        endDate.setHours(0, 0, 0, 0);
        startDate.setHours(0, 0, 0, 0);
        
        // Kiểm tra trạng thái mặc định (bỏ qua nếu "Done")
        if (!conditions.includes('Bỏ qua nếu') && taskStatus === 'Done') {
          Logger.log('Bỏ qua hàng ' + (j + 1) + ': Trạng thái là "Done".');
          continue;
        }
        
        // Kiểm tra điều kiện bỏ qua
        if (shouldSkip(taskRow, skipConditions)) {
          Logger.log('Bỏ qua hàng ' + (j + 1) + ': Không thỏa điều kiện bỏ qua.');
          continue;
        }
        
        // Kiểm tra điều kiện phải thỏa
        if (checkConditions.length > 0 && !shouldCheck(taskRow, checkConditions)) {
          Logger.log('Bỏ qua hàng ' + (j + 1) + ': Không thỏa điều kiện kiểm tra.');
          continue;
        }
        
        // Tính ngày đến hạn
        var daysToEnd = Math.floor((endDate - today) / (1000 * 60 * 60 * 24));
        Logger.log('Hàng ' + (j + 1) + ': Days to end = ' + daysToEnd);
        
        var alertReason = '';
        if (beforeDays > 0 && daysToEnd > 0 && daysToEnd <= beforeDays) {
          alertReason = '⏳ Due soon ( ' + beforeDays + ' ngày)'; //“Due soon” không chỉ báo riêng cho đúng ngày trước deadline, mà bao gồm tất cả ngày trước đó trong khoảng beforeDays
        } else if (daysToEnd === 0) {
          alertReason = '⚠️ Due today';
        } else if (afterDays > 0 && daysToEnd === -afterDays) {
          alertReason = '⛔️ Overdue (' + afterDays + ' ngày)';
        } else if (daysToEnd < 0) {
          alertReason = '⛔️ Overdue';
        }
        
        if (alertReason) {
          //alerts.push(taskName + ': ' + alertReason + ' (End: ' + Utilities.formatDate(endDate, 'GMT+7', 'yyyy-MM-dd') + ', Status: ' + (taskStatus || 'N/A') + ')');
          alerts.push(alertReason + ' (End: ' + Utilities.formatDate(endDate, 'GMT+7', 'yyyy-MM-dd') + ', ' + (taskStatus || 'N/A') + '): ' + taskName);
          Logger.log('Thêm cảnh báo: ' + alerts[alerts.length - 1]);
        }
      }
      
      if (alerts.length > 0) {
        var subject = '🚨 Warning Plan: ' + targetSpreadsheet.getName();
        var body = 'Các task cần chú ý:\n\n' + alerts.join('\n') + '\n\nLink file: ' + targetSpreadsheet.getUrl();
        //var body = 'Các task cần chú ý:\n' 
         //+ alerts.map(task => '- ' + task).join('\n') 
         //+ '\n\nLink file: ' + targetSpreadsheet.getUrl();
        //var body = 'Các task cần chú ý:\n'
        //  + alerts.map(task => '- ' + task).join('\n')
        //  + '\n\nLink file: ' + targetSpreadsheet.getUrl();

        Logger.log('Chuẩn bị gửi cảnh báo: Subject = ' + subject + ', Body = ' + body);
        
        // Gửi qua Email
        if (sendMethod === 'email' || sendMethod === 'both') {
          validEmails.forEach(email => {
            try {
              MailApp.sendEmail(email, subject, body);
              Logger.log('Gửi email thành công tới: ' + email);
            } catch (e) {
              Logger.log('LỖI khi gửi email tới ' + email + ': ' + e.message);
            }
          });
        }
        
        // Gửi qua Telegram
        if (sendMethod === 'telegram' || sendMethod === 'both') {
          validChatIds.forEach(chatId => {
            try {
              var url = 'https://api.telegram.org/bot' + TELEGRAM_BOT_TOKEN + '/sendMessage';
              var payload = {
                chat_id: chatId,
                text: subject + '\n\n' + body
              };
              var options = {
                method: 'post',
                contentType: 'application/json',
                payload: JSON.stringify(payload)
              };
              var response = UrlFetchApp.fetch(url, options);
              Logger.log('Gửi Telegram thành công tới Chat ID: ' + chatId + ', Response: ' + response.getContentText());
            } catch (e) {
              Logger.log('LỖI khi gửi Telegram tới Chat ID ' + chatId + ': ' + e.message);
            }
          });
        }
        
        // Cập nhật trạng thái
        sentCount++;
        var newStatus = 'Đã gửi: ' + sentCount + ' lần, lần cuối: ' + Utilities.formatDate(today, 'GMT+7', 'yyyy-MM-dd');
        configSheet.getRange(i + 1, 15).setValue(newStatus);
        Logger.log('Cập nhật trạng thái hàng ' + (i + 1) + ': ' + newStatus);
      } else {
        Logger.log('Hàng ' + (i + 1) + ': Không có task nào cần cảnh báo.');
      }
    } catch (e) {
      Logger.log('LỖI khi xử lý file ' + fileId + ': ' + e.message);
    }
  }
  Logger.log('Kết thúc checkDeadlines.');
}

// Function to get config data
function getConfigData() {
  var configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  if (!configSheet) {
    return { configData: [], alerts: [] };
  }
  
  var data = configSheet.getDataRange().getValues();
  var configData = data.slice(1); // Bỏ header
  var alerts = [];

  var today = new Date();
  today.setHours(0, 0, 0, 0);

  for (var i = 0; i < configData.length; i++) {
    var row = configData[i];
    var fileId = row[1].trim();
    var projectName = row[2].trim() || '';
    var sheetName = row[3].trim() || null;
    var taskNameColLetter = row[4].trim().toUpperCase();
    var startColLetter = row[5].trim().toUpperCase();
    var endColLetter = row[6].trim().toUpperCase();
    var statusColLetter = row[7].trim().toUpperCase();
    var recipients = row[8].trim().split(',').map(e => e.trim()).join(',');
    var sendMethod = row[9].trim().toLowerCase() || 'email';
    var maxAlerts = parseInt(row[10]) || 0;
    var beforeDays = parseInt(row[11]) || 0;
    var afterDays = parseInt(row[12]) || 0;
    var conditions = row[13].trim() || '';
    var status = row[14].trim() || 'Đã gửi: 0 lần';
    var active = row[15].trim() || 'Yes';

    if (active !== 'Yes') continue; // Skip if not active

    var sentCount = parseInt(status.match(/Đã gửi: (\d+) lần/)?.[1] || 0);
    if (sentCount >= maxAlerts && maxAlerts > 0) continue;

    try {
      var targetSpreadsheet = SpreadsheetApp.openById(fileId);
      var targetSheet = sheetName ? targetSpreadsheet.getSheetByName(sheetName) : targetSpreadsheet.getSheets()[0];
      if (!targetSheet) continue;

      var targetData = targetSheet.getDataRange().getValues();
      var taskNameColIndex = columnLetterToIndex(taskNameColLetter);
      var startColIndex = columnLetterToIndex(startColLetter);
      var endColIndex = columnLetterToIndex(endColLetter);
      var statusColIndex = statusColLetter ? columnLetterToIndex(statusColLetter) : 0;

      var rowRange = parseRowRange(conditions);
      var skipConditions = parseConditions(conditions, 'Bỏ qua nếu');
      var checkConditions = parseConditions(conditions, 'Kiểm tra nếu');

      for (var j = (rowRange[0] - 1); j < Math.min(rowRange[1], targetData.length); j++) {
        var taskRow = targetData[j];
        var taskName = taskNameColIndex ? taskRow[taskNameColIndex - 1] || 'Task ở hàng ' + (j + 1) : 'Task ở hàng ' + (j + 1);
        var startDate = taskRow[startColIndex - 1];
        var endDate = taskRow[endColIndex - 1];
        var taskStatus = statusColIndex ? taskRow[statusColIndex - 1] : '';

        if (!startDate || !endDate || !(startDate instanceof Date) || !(endDate instanceof Date)) continue;
        
        endDate.setHours(0, 0, 0, 0);
        startDate.setHours(0, 0, 0, 0);

        if (!conditions.includes('Bỏ qua nếu') && taskStatus === 'Done') continue;
        if (shouldSkip(taskRow, skipConditions)) continue;
        if (checkConditions.length > 0 && !shouldCheck(taskRow, checkConditions)) continue;

        var daysToEnd = Math.floor((endDate - today) / (1000 * 60 * 60 * 24));
        var alertReason = '';
        if (beforeDays > 0 && daysToEnd === beforeDays) {
          alertReason = 'Sắp đến hạn (trước ' + beforeDays + ' ngày)';
        } else if (daysToEnd === 0) {
          alertReason = 'Hôm nay là ngày hết hạn';
        } else if (afterDays > 0 && daysToEnd === -afterDays) {
          alertReason = 'Đã trễ hạn (sau ' + afterDays + ' ngày)';
        } else if (daysToEnd < 0) {
          alertReason = 'Đã trễ hạn';
        }

        if (alertReason) {
          alerts.push(taskName + ': ' + alertReason + ' (End: ' + Utilities.formatDate(endDate, 'GMT+7', 'yyyy-MM-dd') + ', Status: ' + (taskStatus || 'N/A') + ') - File: ' + targetSpreadsheet.getName());
        }
      }
    } catch (e) {
      Logger.log('LỖI khi xử lý file ' + fileId + ': ' + e.message);
    }
  }

  return { configData: configData, alerts: alerts };
}

// Add new config
function addConfig(data) {
  var configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  var lastRow = configSheet.getLastRow();
  var stt = lastRow > 1 ? parseInt(configSheet.getRange(lastRow, 1).getValue()) + 1 : 1;
  var newRow = [
    stt,
    data.fileId,
    data.projectName,
    data.sheetName,
    data.taskNameCol,
    data.startCol,
    data.endCol,
    data.statusCol,
    data.recipients,
    data.sendMethod,
    data.maxAlerts,
    data.beforeDays,
    data.afterDays,
    data.conditions,
    'Đã gửi: 0 lần',
    'Yes',
    data.link
  ];
  configSheet.appendRow(newRow);
}

// Update config
// Update config
function updateConfig(index, data) {
  var configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  var rowNum = index + 2; // Header + index
  if (!configSheet) {
    Logger.log('LỖI: Sheet "Config" không tồn tại.');
    return;
  }

  // Lấy giá trị hiện tại của hàng để giữ nguyên nếu không thay đổi
  var currentRow = configSheet.getRange(rowNum, 1, 1, configSheet.getLastColumn()).getValues()[0];

  // Cập nhật các cột theo dữ liệu gửi lên
  configSheet.getRange(rowNum, 2).setValue(data.fileId || currentRow[1]); // File ID
  configSheet.getRange(rowNum, 3).setValue(data.projectName || currentRow[2]); // Project Name
  configSheet.getRange(rowNum, 4).setValue(data.sheetName || currentRow[3]); // Sheet Name
  configSheet.getRange(rowNum, 5).setValue(data.taskNameCol || currentRow[4]); // Task Name Col
  configSheet.getRange(rowNum, 6).setValue(data.startCol || currentRow[5]); // Start Date Col
  configSheet.getRange(rowNum, 7).setValue(data.endCol || currentRow[6]); // End Date Col
  configSheet.getRange(rowNum, 8).setValue(data.statusCol || currentRow[7]); // Status Col
  configSheet.getRange(rowNum, 9).setValue(data.recipients || currentRow[8]); // Recipients
  configSheet.getRange(rowNum, 10).setValue(data.sendMethod || currentRow[9]); // Send Method
  configSheet.getRange(rowNum, 11).setValue(data.maxAlerts || currentRow[10]); // Max Alerts
  configSheet.getRange(rowNum, 12).setValue(data.beforeDays || currentRow[11]); // Before Days
  configSheet.getRange(rowNum, 13).setValue(data.afterDays || currentRow[12]); // After Days
  configSheet.getRange(rowNum, 14).setValue(data.conditions || currentRow[13]); // Conditions
  configSheet.getRange(rowNum, 15).setValue(data.status || currentRow[14]); // Status
  configSheet.getRange(rowNum, 16).setValue(data.active || currentRow[15]); // Active (Yes/No)
  configSheet.getRange(rowNum, 17).setValue(data.link || currentRow[16]); // Link

  Logger.log('Cập nhật hàng ' + rowNum + ' thành công: Active = ' + data.active + ', Link = ' + data.link);
}
// Toggle active
function toggleActive(index) {
  var configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  var rowNum = index + 2;
  var currentActive = configSheet.getRange(rowNum, 16).getValue();
  configSheet.getRange(rowNum, 16).setValue(currentActive === 'Yes' ? 'No' : 'Yes');
}

// Helper functions
function columnLetterToIndex(letter) {
  return letter ? letter.charCodeAt(0) - 64 : 0;
}

function parseRowRange(conditions) {
  var match = conditions.match(/hàng từ (\d+) đến (\d+)/);
  return match ? [parseInt(match[1]), parseInt(match[2])] : [2, 9999];
}

function parseConditions(conditions, prefix) {
  var regex = new RegExp(prefix + ' cột (\\w+) = "([^"]+)"', 'gi');
  var conds = [];
  var match;
  while ((match = regex.exec(conditions)) !== null) {
    conds.push({col: match[1].toUpperCase(), val: match[2]});
  }
  return conds;
}

function shouldSkip(row, skipConditions) {
  for (var cond of skipConditions) {
    var colIndex = columnLetterToIndex(cond.col);
    if (colIndex && row[colIndex - 1] === cond.val) return true;
  }
  return false;
}

function shouldCheck(row, checkConditions) {
  for (var cond of checkConditions) {
    var colIndex = columnLetterToIndex(cond.col);
    if (colIndex && row[colIndex - 1] === cond.val) return true;
  }
  return false;
}
function checkAlerts(index) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Config");
  var data = sheet.getDataRange().getValues();
  var configRow = data[index + 1]; // +1 because index 0 is header
  var fileId = configRow[1]; // File ID
  var sheetName = configRow[3]; // Sheet Name
  var endCol = configRow[6].charCodeAt(0) - 64; // Convert column letter to index (e.g., 'G' -> 7)
  var taskNameCol = configRow[4].charCodeAt(0) - 64; // Task Name column
  var targetSheet = SpreadsheetApp.openById(fileId).getSheetByName(sheetName);
  var taskData = targetSheet.getDataRange().getValues();
  var alerts = [];
  for (var i = 1; i < taskData.length; i++) {
    var endDate = new Date(taskData[i][endCol - 1]); // -1 because column index is 1-based
    var today = new Date();
    var diffDays = Math.ceil((endDate - today) / (1000 * 60 * 60 * 24));
    if (diffDays <= 1 && diffDays >= 0) {
      alerts.push(taskData[i][taskNameCol - 1] + ' Sắp đến hạn (trước 1 ngày)');
    } else if (diffDays < 0) {
      alerts.push(taskData[i][taskNameCol - 1] + ' Đã trễ hạn');
    }
  }
  return alerts;
}
