// ===== Apps Script v4 — ระบบบันทึกงานจัดซื้อจัดจ้าง =====
// วางทับโค้ดเดิมทั้งหมดใน Code.gs แล้วกด Save
// จากนั้น Deploy > Manage deployments > ดินสอ > Version: New version > Deploy

var CB_ = '';
function out_(obj){
  var json = JSON.stringify(obj);
  if(CB_){ // JSONP: ใช้เมื่อเว็บส่ง callback มา (สำรองกรณี fetch ถูกบล็อก)
    return ContentService.createTextOutput(CB_ + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function cell_(v){
  if(v instanceof Date){
    var d = ('0'+v.getDate()).slice(-2), m = ('0'+(v.getMonth()+1)).slice(-2);
    var y = v.getFullYear();
    if(y < 2400) y += 543;            // ปี ค.ศ. -> พ.ศ. (ถ้าเก็บเป็น พ.ศ. อยู่แล้วไม่บวกซ้ำ)
    return d + '/' + m + '/' + y;
  }
  return v;
}

// คอลัมน์ที่ต้องเก็บเป็น "ข้อความ" เพื่อกัน Google Sheets แปลงเป็นวันที่เอง (เช่น 05/10/2569 ถูกสลับวัน/เดือน, 69-06 กลายเป็นวันที่)
var TEXT_COLS_ = {
  'ทะเบียนจัดซื้อจัดจ้าง': [8, 9, 10, 11, 14, 18],
  'ทะเบียนผู้ขาย': [4],
  'รายการงบประมาณ': [11]
};

function writeRow_(sh, rowNum, arr){
  (TEXT_COLS_[sh.getName()] || []).forEach(function(c){
    if(c <= arr.length) sh.getRange(rowNum, c).setNumberFormat('@');
  });
  var clean = arr.map(function(v){ return v === null || v === undefined ? '' : v; });
  sh.getRange(rowNum, 1, 1, clean.length).setValues([clean]);
}

function doGet(e){
  try{
    var p = e.parameter || {};
    CB_ = /^[A-Za-z0-9_$.]+$/.test(p.callback || '') ? p.callback : '';
    var action = p.action || 'list';
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    if(action === 'setup_vendors'){
      var vs = ss.getSheetByName('ทะเบียนผู้ขาย');
      if(!vs){
        vs = ss.insertSheet('ทะเบียนผู้ขาย');
        vs.getRange(1,1,1,5).setValues([['ชื่อบริษัท/ร้านค้า','ที่อยู่','ชื่อผู้ติดต่อ','เบอร์ติดต่อ','หมายเหตุ']]);
        vs.getRange(1,1,1,5).setFontWeight('bold');
      }
      return out_({success:true});
    }

    if(action === 'setup_register_cols'){
      var rs = ss.getSheetByName('ทะเบียนจัดซื้อจัดจ้าง');
      if(rs){
        if(!rs.getRange(3,17).getValue()) rs.getRange(3,17).setValue('ปีงบประมาณ');
        if(!rs.getRange(3,18).getValue()) rs.getRange(3,18).setValue('รหัสงบประมาณ');
      }
      return out_({success:true});
    }

    if(action === 'setup_budget_code'){
      var bs = ss.getSheetByName('รายการงบประมาณ');
      if(bs && !bs.getRange(3,11).getValue()){
        bs.getRange(3,11).setValue('รหัสงบประมาณ');
      }
      return out_({success:true});
    }

    var sh = ss.getSheetByName(p.sheet);
    if(!sh) return out_({error:'ไม่พบชีตชื่อ: ' + p.sheet});

    if(action === 'list'){
      var lr = sh.getLastRow(), lc = Math.max(sh.getLastColumn(), 18);
      var rows = lr ? sh.getRange(1,1,lr,lc).getValues().map(function(r){ return r.map(cell_); }) : [];
      return out_({rows:rows});
    }

    if(action === 'add'){
      var row = JSON.parse(p.row);
      writeRow_(sh, sh.getLastRow() + 1, row);
      return out_({success:true});
    }

    if(action === 'update'){
      var n = parseInt(p.row_num, 10);
      var arr = JSON.parse(p.row);
      if(!n || n < 1) return out_({error:'row_num ไม่ถูกต้อง'});
      writeRow_(sh, n, arr);
      return out_({success:true});
    }

    if(action === 'delete'){
      var d = parseInt(p.row_num, 10);
      if(!d || d < 2) return out_({error:'row_num ไม่ถูกต้อง'});
      sh.deleteRow(d);
      return out_({success:true});
    }

    return out_({error:'ไม่รู้จัก action: ' + action});
  }catch(err){
    return out_({error:String(err)});
  }
}


// =====================================================================
// แจ้งเตือนสัญญา/รายการใกล้หมดอายุผ่าน Telegram (v5)
// ส่งทุกวัน จนกว่ารายการนั้นจะมีสถานะ "เสร็จ / ไม่ต่อ / ยกเลิก / ปิดงาน"
// ตั้งค่า Token และ Chat ID ที่ Project Settings > Script properties
// (ห้ามใส่ Token ในโค้ดนี้ เพราะไฟล์นี้อยู่บน GitHub แบบ Public)
//   TELEGRAM_TOKEN   = โทเค็นจาก @BotFather
//   TELEGRAM_CHAT_ID = เลข Chat ID ของคุณ (หรือของกลุ่ม)
// =====================================================================

var NOTIFY_DAYS_ = 90;                 // แจ้งเมื่อเหลือไม่เกินกี่วัน
var REGISTER_SHEET_ = 'ทะเบียนจัดซื้อจัดจ้าง';
var CLOSED_RE_ = /เสร็จ|ไม่ต่อ|ยกเลิก|ปิดงาน/;

function parseThaiDate_(v){
  if(v instanceof Date){
    var y = v.getFullYear();
    return new Date(y > 2400 ? y - 543 : y, v.getMonth(), v.getDate());
  }
  var parts = String(v || '').trim().split('/');
  if(parts.length !== 3) return null;
  var d = parseInt(parts[0],10), m = parseInt(parts[1],10), y = parseInt(parts[2],10);
  if(!d || !m || !y) return null;
  return new Date(y > 2400 ? y - 543 : y, m - 1, d);
}

function sendTelegram_(text){
  var props = PropertiesService.getScriptProperties();
  var token = props.getProperty('TELEGRAM_TOKEN');
  var chat = props.getProperty('TELEGRAM_CHAT_ID');
  if(!token || !chat) throw new Error('ยังไม่ได้ตั้ง TELEGRAM_TOKEN / TELEGRAM_CHAT_ID ใน Script properties');
  var res = UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/sendMessage', {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({chat_id: chat, text: text}),
    muteHttpExceptions: true
  });
  if(res.getResponseCode() !== 200) throw new Error('Telegram ตอบกลับ ' + res.getResponseCode() + ': ' + res.getContentText());
}

// ส่งข้อความทดสอบ — เลือกฟังก์ชันนี้แล้วกด Run หนึ่งครั้งเพื่อเช็คว่าตั้งค่าถูก
function testTelegram(){
  sendTelegram_('✅ ทดสอบแจ้งเตือนจากระบบจัดซื้อจัดจ้าง แผนกสารสนเทศ — เชื่อมต่อ Telegram สำเร็จ');
}

// ตรวจรายการและส่งสรุป (ฟังก์ชันนี้ถูกเรียกทุกวันโดย Trigger)
function checkExpiry(){
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(REGISTER_SHEET_);
  if(!sh) throw new Error('ไม่พบชีต ' + REGISTER_SHEET_);
  var lr = sh.getLastRow();
  if(lr < 4) return;
  var rows = sh.getRange(4, 1, lr - 3, 16).getValues();
  var today = new Date(); today.setHours(0,0,0,0);
  var expired = [], soon = [];
  rows.forEach(function(r){
    var name = r[2], status = String(r[14] || '');
    if(!name || CLOSED_RE_.test(status)) return;
    var dt = parseThaiDate_(r[10]);
    if(!dt) return;
    var days = Math.round((dt - today) / 86400000);
    var line = '• ' + name + '\n   หมดอายุ ' + cellText_(r[10]) +
               (days < 0 ? ' (เกินมาแล้ว ' + (-days) + ' วัน)' : ' (อีก ' + days + ' วัน)') +
               '\n   ผู้ขาย: ' + (r[11] || '-') + ' | ผู้รับผิดชอบ: ' + (r[12] || '-');
    if(days < 0) expired.push({d: days, t: line});
    else if(days <= NOTIFY_DAYS_) soon.push({d: days, t: line});
  });
  if(!expired.length && !soon.length) return;       // ไม่มีอะไรต้องแจ้ง
  soon.sort(function(a,b){ return a.d - b.d; });
  expired.sort(function(a,b){ return b.d - a.d; });
  var msg = '🔔 แจ้งเตือนสัญญา/รายการใกล้หมดอายุ\n';
  if(soon.length) msg += '\n⏳ ใกล้หมดอายุ (≤ ' + NOTIFY_DAYS_ + ' วัน) ' + soon.length + ' รายการ\n' + soon.map(function(x){return x.t;}).join('\n');
  if(expired.length) msg += '\n\n❌ หมดอายุแล้ว ' + expired.length + ' รายการ\n' + expired.map(function(x){return x.t;}).join('\n');
  msg += '\n\nแจ้งต่อเนื่องจนกว่าจะปิดงาน (กด "เสร็จแล้ว" หรือ "ไม่ต่ออายุ" ในเว็บ)';
  // Telegram จำกัด 4096 ตัวอักษรต่อข้อความ → แบ่งส่ง
  for(var i = 0; i < msg.length; i += 3800) sendTelegram_(msg.substring(i, i + 3800));
}

function cellText_(v){
  if(v instanceof Date) return cell_(v);
  return String(v);
}

// รันครั้งเดียวเพื่อตั้งเวลาส่งอัตโนมัติทุกวัน 08:00 (ตามเขตเวลาของโปรเจกต์)
function createDailyTrigger(){
  ScriptApp.getProjectTriggers().forEach(function(t){
    if(t.getHandlerFunction() === 'checkExpiry') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('checkExpiry').timeBased().everyDays(1).atHour(8).create();
}
