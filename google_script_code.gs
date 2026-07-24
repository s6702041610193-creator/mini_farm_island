// ==========================================
// GOOGLE APPS SCRIPT BACKEND CODE (Code.gs)
// Sheet ID: 1SvjrbVQiXKv1kCUbX6pBzijYMchQfDzULd-mvmB89M9zfEhIZ6gAvPvv
// ==========================================

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);
  
  try {
    var sheetId = '1SvjrbVQiXKv1kCUbX6pBzijYMchQfDzULd-mvmB89M9zfEhIZ6gAvPvv';
    var ss = SpreadsheetApp.openById(sheetId);
    
    // Ensure required sheets exist
    var usersSheet = getOrCreateSheet(ss, 'Users', ['Timestamp', 'Username', 'Password', 'Name']);
    var logsSheet = getOrCreateSheet(ss, 'Logs', ['Timestamp', 'Username', 'Action']);
    var savesSheet = getOrCreateSheet(ss, 'Saves', ['Timestamp', 'Username', 'SaveData']);
    
    var data = JSON.parse(e.postData.contents);
    var action = data.action;
    
    if (action === 'register') {
      var username = data.username;
      var password = data.password;
      var name = data.name;
      
      // Check existing user
      var users = usersSheet.getDataRange().getValues();
      for (var i = 1; i < users.length; i++) {
        if (users[i][1] === username) {
          return responseJSON({ status: 'error', message: 'ชื่อผู้ใช้นี้ถูกใช้งานแล้ว' });
        }
      }
      
      usersSheet.appendRow([new Date(), username, password, name]);
      logsSheet.appendRow([new Date(), username, 'REGISTER']);
      return responseJSON({ status: 'success', message: 'สมัครสมาชิกสำเร็จ' });
    }
    
    else if (action === 'login') {
      var username = data.username;
      var password = data.password;
      
      var users = usersSheet.getDataRange().getValues();
      var found = false;
      for (var i = 1; i < users.length; i++) {
        if (users[i][1] === username && users[i][2] === password) {
          found = true;
          break;
        }
      }
      
      if (found) {
        logsSheet.appendRow([new Date(), username, 'LOGIN']);
        return responseJSON({ status: 'success', message: 'เข้าสู่ระบบสำเร็จ' });
      } else {
        return responseJSON({ status: 'error', message: 'ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง' });
      }
    }
    
    else if (action === 'log_login') {
      logsSheet.appendRow([new Date(), data.username, 'LOGIN_ACTIVITY']);
      return responseJSON({ status: 'success' });
    }
    
    else if (action === 'save_game') {
      var username = data.username;
      var saveData = data.saveData;
      
      var saves = savesSheet.getDataRange().getValues();
      var rowIndex = -1;
      for (var i = 1; i < saves.length; i++) {
        if (saves[i][1] === username) {
          rowIndex = i + 1;
          break;
        }
      }
      
      if (rowIndex !== -1) {
        savesSheet.getRange(rowIndex, 1).setValue(new Date());
        savesSheet.getRange(rowIndex, 3).setValue(saveData);
      } else {
        savesSheet.appendRow([new Date(), username, saveData]);
      }
      
      logsSheet.appendRow([new Date(), username, 'SAVE_GAME']);
      return responseJSON({ status: 'success', message: 'บันทึกข้อมูลเรียบร้อย' });
    }
    
    else if (action === 'load_game') {
      var username = data.username;
      var saves = savesSheet.getDataRange().getValues();
      
      for (var i = 1; i < saves.length; i++) {
        if (saves[i][1] === username) {
          return responseJSON({ status: 'success', saveData: saves[i][2] });
        }
      }
      return responseJSON({ status: 'error', message: 'ไม่พบข้อมูลเซฟ' });
    }
    
    return responseJSON({ status: 'error', message: 'Unknown action' });
    
  } catch (err) {
    return responseJSON({ status: 'error', message: err.toString() });
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  return ContentService.createTextOutput("Mini Farm Island API Active");
}

function getOrCreateSheet(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
  }
  return sheet;
}

function responseJSON(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
