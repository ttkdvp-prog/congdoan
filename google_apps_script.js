/**
 * ============================================================================
 * CÔNG ĐOÀN SỐ - GOOGLE APPS SCRIPT BACKEND (Code.gs)
 * Spreadsheet ID: 1tkGGIDQbiJj45MCdWXYFrA66u7ovWjK5EHvVLyLfGFc
 * 
 * Hướng dẫn 3 bước ngắn gọn:
 * 1. Trên Google Sheet "He_thong_Cong_doan_so" (link bên dưới):
 *    https://docs.google.com/spreadsheets/d/1tkGGIDQbiJj45MCdWXYFrA66u7ovWjK5EHvVLyLfGFc/edit
 * 2. Vào menu "Tiện ích mở rộng" -> "Apps Script", dán toàn bộ mã này vào Code.gs và bấm Save (Lưu).
 * 3. Bấm nút "Triển khai" (Deploy) ở góc trên bên phải -> "Tạo bản triển khai mới" (New deployment):
 *    - Loại: Ứng dụng web (Web App)
 *    - Thực thi dưới danh nghĩa: Tôi (Me)
 *    - Ai có quyền truy cập: Bất kỳ ai (Anyone)
 *    - Bấm Triển khai, cấp quyền và copy Web App URL thu được (dạng https://script.google.com/macros/s/.../exec).
 * ============================================================================
 */

var SPREADSHEET_ID = "1tkGGIDQbiJj45MCdWXYFrA66u7ovWjK5EHvVLyLfGFc";

function getSpreadsheet() {
  try {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  } catch (e) {
    return SpreadsheetApp.getActiveSpreadsheet();
  }
}

function doGet(e) {
  var action = e && e.parameter && e.parameter.action ? e.parameter.action : 'getData';
  
  if (action === 'getData') {
    return handleGetData();
  }
  
  return respondJSON({ status: 'ok', message: 'API Cong Doan So Active', spreadsheetId: SPREADSHEET_ID });
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var action = data.action;
    
    if (action === 'create') {
      return handleCreate(data.payload);
    } else if (action === 'update') {
      return handleUpdate(data.payload);
    } else if (action === 'approve') {
      return handleApprove(data.payload);
    }
    
    return respondJSON({ status: 'error', message: 'Action không hợp lệ: ' + action });
  } catch (err) {
    return respondJSON({ status: 'error', message: err.toString() });
  }
}

function handleGetData() {
  var ss = getSpreadsheet();
  
  var result = {
    status: 'success',
    timestamp: new Date().toISOString(),
    DL_CHAM_LO: readSheetData(ss, 'DL_CHAM_LO'),
    TIEP_NHAN: readSheetData(ss, 'TIEP_NHAN'),
    QUY_CHE: readSheetData(ss, 'QUY_CHE')
  };
  
  return respondJSON(result);
}

function readSheetData(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  
  var headers = values[0];
  var rows = [];
  
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var obj = {};
    var isEmpty = true;
    for (var j = 0; j < headers.length; j++) {
      var headerKey = String(headers[j]).trim();
      if (headerKey) {
        var val = row[j];
        if (val instanceof Date) {
          val = Utilities.formatDate(val, Session.getScriptTimeZone(), "dd/MM/yyyy");
        }
        if (val !== "" && val !== null && val !== undefined) isEmpty = false;
        obj[headerKey] = val;
      }
    }
    if (!isEmpty) rows.push(obj);
  }
  return rows;
}

function handleCreate(payload) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('TIEP_NHAN') || ss.getSheetByName('DL_CHAM_LO');
  
  if (!sheet) {
    return respondJSON({ status: 'error', message: 'Không tìm thấy sheet TIEP_NHAN' });
  }
  
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var newRow = [];
  
  for (var i = 0; i < headers.length; i++) {
    var key = String(headers[i]).trim();
    var val = payload[key] !== undefined ? payload[key] : (payload[mapKey(key)] !== undefined ? payload[mapKey(key)] : '');
    newRow.push(val);
  }
  
  sheet.appendRow(newRow);
  return respondJSON({ status: 'success', message: 'Đã lưu tự động vào Google Sheet thành công', ma: payload.ma });
}

function handleUpdate(payload) {
  var ss = getSpreadsheet();
  var sheets = ['TIEP_NHAN', 'DL_CHAM_LO'];
  var updated = false;
  
  for (var s = 0; s < sheets.length; s++) {
    var sheet = ss.getSheetByName(sheets[s]);
    if (!sheet) continue;
    
    var values = sheet.getDataRange().getValues();
    if (values.length < 2) continue;
    
    var headers = values[0];
    var maColIdx = -1;
    
    for (var h = 0; h < headers.length; h++) {
      var headStr = String(headers[h]).trim().toLowerCase();
      if (headStr === 'mã' || headStr === 'mã hồ sơ' || headStr === 'ma') {
        maColIdx = h;
        break;
      }
    }
    
    if (maColIdx === -1) continue;
    
    for (var r = 1; r < values.length; r++) {
      if (String(values[r][maColIdx]).trim() === String(payload.ma).trim()) {
        for (var c = 0; c < headers.length; c++) {
          var key = String(headers[c]).trim();
          var val = payload[key] !== undefined ? payload[key] : payload[mapKey(key)];
          if (val !== undefined) {
            sheet.getRange(r + 1, c + 1).setValue(val);
          }
        }
        updated = true;
        break;
      }
    }
    if (updated) break;
  }
  
  return respondJSON({ status: updated ? 'success' : 'not_found', message: updated ? 'Đã cập nhật dòng trong Google Sheet' : 'Không tìm thấy dòng tương ứng' });
}

function handleApprove(payload) {
  payload.trangThai = 'Đã hoàn thành';
  payload.canhBao = 'Đúng hạn';
  return handleUpdate(payload);
}

function mapKey(header) {
  var mapping = {
    'Mã hồ sơ': 'ma', 'Mã': 'ma',
    'Nội dung': 'noiDung', 'Nội dung vụ việc': 'noiDung',
    'Loại chăm lo': 'loai', 'Loại': 'loai',
    'Kinh phí': 'soTien', 'Mức đề nghị': 'soTien', 'Số tiền': 'soTien',
    'Địa bàn': 'diaBan',
    'Đoàn viên': 'doanVien', 'Người được hỗ trợ': 'nguoiHoTro',
    'Trạng thái': 'trangThai',
    'Cảnh báo': 'canhBao',
    'Ngày': 'ngayStr'
  };
  return mapping[header] || header;
}

function respondJSON(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
