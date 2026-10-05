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
    return d + '/' + m + '/' + (v.getFullYear()+543);
  }
  return v;
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
      var lr = sh.getLastRow(), lc = Math.max(sh.getLastColumn(), 11);
      var rows = lr ? sh.getRange(1,1,lr,lc).getValues().map(function(r){ return r.map(cell_); }) : [];
      return out_({rows:rows});
    }

    if(action === 'add'){
      var row = JSON.parse(p.row);
      sh.appendRow(row);
      return out_({success:true});
    }

    if(action === 'update'){
      var n = parseInt(p.row_num, 10);
      var arr = JSON.parse(p.row);
      if(!n || n < 1) return out_({error:'row_num ไม่ถูกต้อง'});
      sh.getRange(n, 1, 1, arr.length).setValues([arr]);
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
