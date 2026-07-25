/**
 * ============================================================================
 * CÔNG ĐOÀN SỐ - GOOGLE APPS SCRIPT BACKEND (Code.gs)
 * Spreadsheet ID: 1tkGGIDQbiJj45MCdWXYFrA66u7ovWjK5EHvVLyLfGFc
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
  var callback = e && e.parameter && e.parameter.callback ? e.parameter.callback : (e && e.parameter && e.parameter.prefix ? e.parameter.prefix : null);
  
  var data;
  if (action === 'getData') {
    data = handleGetDataRaw();
  } else {
    data = { status: 'ok', message: 'API Cong Doan So Active', spreadsheetId: SPREADSHEET_ID };
  }
  
  return respondOutput(data, callback);
}

function handleGetDataRaw() {
  var ss = getSpreadsheet();
  return {
    status: 'success',
    timestamp: new Date().toISOString(),
    DL_CHAM_LO: readSheetData(ss, 'DL_CHAM_LO'),
    TIEP_NHAN: readSheetData(ss, 'TIEP_NHAN'),
    QUY_CHE: readSheetData(ss, 'QUY_CHE')
  };
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
    
    return respondOutput({ status: 'error', message: 'Action không hợp lệ: ' + action }, null);
  } catch (err) {
    return respondOutput({ status: 'error', message: err.toString() }, null);
  }
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
  
  // 1. Write to TIEP_NHAN sheet
  var sheetTiepNhan = ss.getSheetByName('TIEP_NHAN');
  if (sheetTiepNhan) {
    appendRowToSheet(sheetTiepNhan, payload);
  }
  
  // 2. If status is "Đã hoàn thành", write to DL_CHAM_LO sheet as well
  if (payload.trangThai === 'Đã hoàn thành') {
    var sheetChamLo = ss.getSheetByName('DL_CHAM_LO');
    if (sheetChamLo) {
      appendRowToSheet(sheetChamLo, payload);
    }
  }
  
  return respondOutput({ status: 'success', message: 'Đã lưu tự động vào Google Sheet', ma: payload.ma }, null);
}

function handleUpdate(payload) {
  var ss = getSpreadsheet();
  var sheets = ['TIEP_NHAN', 'DL_CHAM_LO'];
  var updated = false;
  
  for (var s = 0; s < sheets.length; s++) {
    var sheet = ss.getSheetByName(sheets[s]);
    if (!sheet) continue;
    
    var res = updateRowInSheet(sheet, payload);
    if (res) updated = true;
  }
  
  // If status is "Đã hoàn thành", ensure it exists in DL_CHAM_LO
  if (payload.trangThai === 'Đã hoàn thành') {
    var sheetChamLo = ss.getSheetByName('DL_CHAM_LO');
    if (sheetChamLo) {
      var existsInChamLo = findRowIndexByMa(sheetChamLo, payload.ma);
      if (existsInChamLo === -1) {
        appendRowToSheet(sheetChamLo, payload);
        updated = true;
      }
    }
  }
  
  return respondOutput({ status: updated ? 'success' : 'not_found', message: updated ? 'Đã cập nhật Google Sheet' : 'Không tìm thấy dòng tương ứng' }, null);
}

function handleApprove(payload) {
  payload.trangThai = 'Đã hoàn thành';
  payload.canhBao = 'Đúng hạn';
  return handleUpdate(payload);
}

function findRowIndexByMa(sheet, ma) {
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return -1;
  
  var headers = values[0];
  var maColIdx = -1;
  
  for (var h = 0; h < headers.length; h++) {
    var headStr = String(headers[h]).trim().toLowerCase();
    if (headStr === 'mã' || headStr === 'mã hồ sơ' || headStr === 'ma' || headStr.indexOf('mã') !== -1) {
      maColIdx = h;
      break;
    }
  }
  
  if (maColIdx === -1) return -1;
  
  for (var r = 1; r < values.length; r++) {
    if (String(values[r][maColIdx]).trim() === String(ma).trim()) {
      return r + 1; // 1-based row index
    }
  }
  return -1;
}

function updateRowInSheet(sheet, payload) {
  var rowIdx = findRowIndexByMa(sheet, payload.ma);
  if (rowIdx === -1) return false;
  
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  for (var c = 0; c < headers.length; c++) {
    var header = String(headers[c]).trim();
    if (!header) continue;
    
    var val = getPayloadVal(payload, header);
    if (val !== undefined && val !== null) {
      sheet.getRange(rowIdx, c + 1).setValue(val);
    }
  }
  return true;
}

function appendRowToSheet(sheet, payload) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var newRow = [];
  
  for (var i = 0; i < headers.length; i++) {
    var header = String(headers[i]).trim();
    var val = getPayloadVal(payload, header);
    newRow.push(val !== undefined && val !== null ? val : '');
  }
  
  sheet.appendRow(newRow);
}

function getPayloadVal(payload, header) {
  var hLower = header.toLowerCase().replace(/[^a-z0-9]/g, '');
  
  // Exact match
  if (payload[header] !== undefined) return payload[header];
  
  // Map by known headers
  if (hLower.indexOf('ma') !== -1) return payload.ma;
  if (hLower.indexOf('ngay') !== -1) {
    if (hLower.indexOf('hoanthanh') !== -1 || hLower.indexOf('chi') !== -1 || hLower.indexOf('duyet') !== -1) {
      return payload.ngayStr || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
    }
    return payload.ngayStr || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
  }
  if (hLower.indexOf('nam') !== -1) return payload.nam || new Date().getFullYear();
  if (hLower.indexOf('thang') !== -1) return payload.thang || (new Date().getMonth() + 1);
  if (hLower.indexOf('noidung') !== -1 || hLower.indexOf('suviec') !== -1 || hLower.indexOf('lydo') !== -1) return payload.noiDung || payload.lyDo;
  if (hLower.indexOf('loai') !== -1) return payload.loai;
  if (hLower.indexOf('tien') !== -1 || hLower.indexOf('kinhphi') !== -1 || hLower.indexOf('muc') !== -1) return payload.soTien;
  if (hLower.indexOf('nguoidenghi') !== -1 || hLower.indexOf('nguoithuchien') !== -1) return payload.nguoiDeNghi || payload.nguoiThucHien;
  if (hLower.indexOf('diaban') !== -1 || hLower.indexOf('to') !== -1) return payload.diaBan;
  if (hLower.indexOf('hotro') !== -1) return payload.nguoiHoTro || payload.doanVien;
  if (hLower.indexOf('doanvien') !== -1) return payload.doanVien || payload.nguoiHoTro;
  if (hLower.indexOf('quanhe') !== -1) return payload.quanHe;
  if (hLower.indexOf('trangthai') !== -1) return payload.trangThai;
  if (hLower.indexOf('canhbao') !== -1) return payload.canhBao;
  if (hLower.indexOf('donvi') !== -1) return payload.donVi;
  
  return '';
}

function respondOutput(data, callback) {
  var jsonStr = JSON.stringify(data);
  if (callback) {
    return ContentService.createTextOutput(callback + "(" + jsonStr + ")")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(jsonStr)
    .setMimeType(ContentService.MimeType.JSON);
}
