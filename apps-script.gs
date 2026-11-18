/**
 * Google Apps Script – nhận đơn hàng từ app Trà Sữa và ghi vào Google Sheets.
 * Cách dùng: mở Google Sheet > Tiện ích mở rộng > Apps Script > dán toàn bộ file này > Lưu
 * > chạy hàm setup một lần > Triển khai > Tùy chọn triển khai mới > Ứng dụng web
 *   (Thực thi với tư cách: Tôi, Người có quyền truy cập: Bất kỳ ai) > copy link /exec vào app.
 */
const SH_ORDERS = 'DonHang';
const SH_ITEMS = 'ChiTiet';
const SH_SUM = 'TongHop';

const H_ORDERS = ['Mã đơn', 'Ngày', 'Giờ', 'Số đơn', 'Các món', 'Số món', 'Tổng tiền', 'Thanh toán', 'Trạng thái', 'Cập nhật lúc'];
const H_ITEMS = ['Mã đơn', 'Ngày', 'Số đơn', 'Món', 'Size', 'SL', 'Đơn giá', 'Topping', 'Ghi chú', 'Thành tiền', 'Trạng thái'];

function setup() {
  const ss = SpreadsheetApp.getActive();
  ensureSheets_(ss);
}

function ensureSheets_(ss) {
  let so = ss.getSheetByName(SH_ORDERS);
  if (!so) {
    so = ss.insertSheet(SH_ORDERS);
    so.getRange(1, 1, 1, H_ORDERS.length).setValues([H_ORDERS]).setFontWeight('bold').setBackground('#f3e6d8');
    so.setFrozenRows(1);
    so.getRange('B:B').setNumberFormat('@'); // giữ ngày dạng yyyy-mm-dd
    so.getRange('C:C').setNumberFormat('@');
    so.getRange('G:G').setNumberFormat('#,##0');
  }
  let si = ss.getSheetByName(SH_ITEMS);
  if (!si) {
    si = ss.insertSheet(SH_ITEMS);
    si.getRange(1, 1, 1, H_ITEMS.length).setValues([H_ITEMS]).setFontWeight('bold').setBackground('#f3e6d8');
    si.setFrozenRows(1);
    si.getRange('B:B').setNumberFormat('@');
    si.getRange('G:G').setNumberFormat('#,##0');
    si.getRange('J:J').setNumberFormat('#,##0');
  }
  let sm = ss.getSheetByName(SH_SUM);
  if (!sm) {
    sm = ss.insertSheet(SH_SUM, 0);
    sm.getRange('A1').setValue('DOANH THU THEO NGÀY').setFontWeight('bold');
    sm.getRange('A2').setFormula(
      "=IFERROR(QUERY(DonHang!B:I,\"select B, count(B), sum(F), sum(G) where I = 'OK' group by B order by B desc label B 'Ngày', count(B) 'Số đơn', sum(F) 'Số món', sum(G) 'Doanh thu'\",1),\"Chưa có dữ liệu\")"
    );
    sm.getRange('F1').setValue('THEO MÓN (TẤT CẢ)').setFontWeight('bold');
    sm.getRange('F2').setFormula(
      "=IFERROR(QUERY(ChiTiet!D:K,\"select D, sum(F), sum(J) where K = 'OK' group by D order by sum(J) desc label D 'Món', sum(F) 'Số món', sum(J) 'Doanh thu'\",1),\"Chưa có dữ liệu\")"
    );
    sm.getRange('J1').setValue('THEO THÁNG').setFontWeight('bold');
    sm.getRange('J2').setFormula(
      "=IFERROR(QUERY({ARRAYFORMULA(LEFT(DonHang!B2:B,7)),DonHang!G2:G,DonHang!I2:I},\"select Col1, count(Col1), sum(Col2) where Col3 = 'OK' and Col1 <> '' group by Col1 order by Col1 desc label Col1 'Tháng', count(Col1) 'Số đơn', sum(Col2) 'Doanh thu'\",0),\"Chưa có dữ liệu\")"
    );
    sm.getRange('D:D').setNumberFormat('#,##0');
    sm.getRange('H:H').setNumberFormat('#,##0');
    sm.getRange('L:L').setNumberFormat('#,##0');
  }
  const def = ss.getSheetByName('Sheet1') || ss.getSheetByName('Trang tính1');
  if (def && def.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(def);
  return { so, si };
}

function doGet() {
  return json_({ ok: true, msg: 'App Trà Sữa – Web App đang chạy. Hãy dán link này vào phần Cài đặt của app.' });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const data = JSON.parse(e.postData.contents || '{}');
    const ss = SpreadsheetApp.getActive();
    const { so, si } = ensureSheets_(ss);

    if (data.action === 'ping') return json_({ ok: true, sheet: ss.getName() });
    if (data.action !== 'upsert') return json_({ ok: false, error: 'action không hợp lệ' });

    // Chỉ mục đơn đã có
    const oLast = so.getLastRow();
    const oIds = oLast > 1 ? so.getRange(2, 1, oLast - 1, 1).getValues().map(r => String(r[0])) : [];
    const oRow = {};
    oIds.forEach((id, i) => (oRow[id] = i + 2));

    let iIds = null; // nạp khi cần
    const newOrders = [], newItems = [], done = [];
    const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

    (data.orders || []).forEach(o => {
      const st = o.status === 'huy' ? 'Đã hủy' : 'OK';
      if (oRow[o.id]) {
        // Đơn đã có: chỉ cập nhật trạng thái
        so.getRange(oRow[o.id], 9, 1, 2).setValues([[st, now]]);
        if (!iIds) {
          const iLast = si.getLastRow();
          iIds = iLast > 1 ? si.getRange(2, 1, iLast - 1, 1).getValues().map(r => String(r[0])) : [];
        }
        iIds.forEach((id, i) => { if (id === o.id) si.getRange(i + 2, 11).setValue(st); });
      } else {
        const items = o.items || [];
        const text = items.map(i => `${i.qty}x ${i.name}${i.size ? ' (' + i.size + ')' : ''}${(i.toppings || []).length ? ' + ' + i.toppings.join(', ') : ''}`).join('; ');
        const cups = items.reduce((s, i) => s + (Number(i.qty) || 0), 0);
        newOrders.push([o.id, o.date, o.time, o.no, text, cups, o.total, o.pay === 'CK' ? 'Chuyển khoản' : 'Tiền mặt', st, now]);
        items.forEach(i => newItems.push([o.id, o.date, o.no, i.name, i.size || '', i.qty, i.unit, (i.toppings || []).join(', '), i.note || '', i.line, st]));
        oRow[o.id] = -1; // tránh trùng trong cùng lô
      }
      done.push(o.id);
    });

    if (newOrders.length) so.getRange(so.getLastRow() + 1, 1, newOrders.length, H_ORDERS.length).setValues(newOrders);
    if (newItems.length) si.getRange(si.getLastRow() + 1, 1, newItems.length, H_ITEMS.length).setValues(newItems);

    return json_({ ok: true, ids: done });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
