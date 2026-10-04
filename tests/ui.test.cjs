'use strict';
// Test giao diện: nạp Index.html vào Chromium headless, google.script.run được nối thẳng vào Code.gs
// chạy trên Google Sheets giả lập. Cần Playwright (npm i -g playwright). Chạy: node tests/ui.test.cjs
const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

function loadPlaywright() {
  try { return require('playwright'); } catch (e) {
    return require(path.join(execSync('npm root -g').toString().trim(), 'playwright'));
  }
}
const { chromium } = loadPlaywright();
const { buildMockEnv } = require('./mock_gas.cjs');

const env = buildMockEnv();
const mod = require('./load_code.cjs')();

const XSS_IMG = '<img src=x onerror="window.__xss=1">';
const XSS_SCRIPT = '<script>window.__xss=2</script>';
const QUOTE_HD = 'HD"002';

const SS = (k) => global.SpreadsheetApp.openById(mod.SS_IDS[k]);
SS('HDMB')._addSheet('HD_RUNG', [
  ['SoHopDong', 'MaRung', 'HoVaTenChuRung', 'ID_RUNG', 'SoCCCD'],
  ['HD001', 'R001', XSS_IMG, 'GPS001', '201'],
  [QUOTE_HD, 'R002', 'Trần Thị B', 'GPS002', '202'],
  ['HD003', 'R003', 'Lê Văn C', 'GPS003', '203'],
]);
SS('HDMB')._addSheet('BaoCao_KiemTra', [['Số HĐ', 'Kết quả', 'Hồ sơ còn thiếu', 'Cảnh báo'], ['HD001', 'Đầy đủ', XSS_SCRIPT, ''], [QUOTE_HD, 'Thiếu', 'CCCD', XSS_IMG]]);
SS('HDMB')._addSheet('HD_Picture', [['ID_HD', 'Picture1'], ['R002', 'javascript:alert(1)']]);
SS('HDMB')._addSheet('HD_GPS', [['ID_KEY_GPS', 'Latitude_Vĩ độ', 'Longtitude_Kinh độ', 'Địa điểm tra GPS'], ['GPS002', '15.9"><img src=x onerror=window.__xss=3>', '108.2', XSS_IMG]]);
SS('HDMB')._addSheet('HD_STK', [['Số hợp đồng', 'Họ và tên người được ủy quyền', 'Số TK nhận tiền', 'Ngân hàng', 'Ủy quyền thanh toán']]);
SS('HDMB')._addSheet('HD_NCC', [['SoHopDong'], ['HD001']]);
const hk = [['STT', 'Ngày nhập', 'Mã hợp đồng', 'SỐ PHIẾU CÂN', 'Số BKLS', 'Họ và tên chủ rừng', 'ĐỊA CHỈ RỪNG', 'Khối lượng (Tấn)', 'Đơn giá', 'Thanh toán', 'Nguồn gốc']];
for (let i = 1; i <= 5; i++) hk.push([i, new Date(2026, 7, i), 'HD001', 'PC' + i, 'BKLS' + i, i === 1 ? XSS_IMG : 'CR' + i, 'ĐC', 10, 1500000, 'Đã TT', 'DT']);
SS('HOSOKEO')._addSheet('HoSoKeo_DN', hk);
SS('HOSOKEO')._addSheet('HoSoRung_DN', [['Mã hợp đồng'], ['HD001']]);
SS('HOSOKEO')._addSheet('ToaDoRung_DN', [['Số hợp đồng'], ['HD001']]);
SS('XUATHANG')._addSheet('NL_PC_XH', [['STT', 'Số phiếu', 'Ngày giờ cân 1', 'Số BKLS', 'Khối lượng (Tấn)', 'Khối lượng (M3)', 'Đơn vị vận chuyển', 'SỐ TKHQ', 'TÀU XUẤT'], [1, 'XH1', new Date(2026, 7, 3), 'BKLS1', 8, 12, 'VT', XSS_IMG, 'Tàu']]);
SS('XUATHANG')._addSheet('NL_DH_XB', [['STT', XSS_IMG], [1, 'ABC']]);
SS('PHIEUCAN')._addSheet('PhieuCan_DN', [['STT'], [1]]);
SS('KHAOSAT_FORM')._addSheet('KhaoSat_FSC', [['Timestamp'], ['2026-08-01']]);
mod.initOwnSheets();
// Bản ghi có cột ngày (Sheets lưu thành kiểu Date) — trước đây làm danh sách trả về null.
mod.addTieuChiThamDinh({ SoHopDong: 'HD003', NgayThamDinh: '2026-10-04', KetLuanThamDinh: 'Đạt' });
mod.updateTienDo('0. Chuẩn bị', '2026-09-01', 'Đúng tiến độ', '');

let pass = 0, fail = 0;
const check = (n, c, d) => { if (c) { pass++; console.log('✅ ' + n); } else { fail++; console.log('❌ ' + n + (d ? ' — ' + d : '')); } };
const calls = [];
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

(async () => {
  const browser = await chromium.launch();
  const page = await browser.newPage();
  const errors = [];
  page.on('pageerror', (e) => errors.push('pageerror: ' + e.message));
  page.on('console', (m) => { if (m.type() === 'error' && !/Failed to load resource|fonts/.test(m.text())) errors.push('console: ' + m.text()); });
  page.on('dialog', (d) => d.dismiss());
  page.on('popup', (p) => p.close().catch(() => {}));
  await page.route(/fonts\.(googleapis|gstatic)\.com/, (r) => r.abort());
  await page.exposeFunction('__gas', (fn, args) => {
    calls.push([fn, args]);
    if (typeof mod[fn] !== 'function') return { __throw: 'Script function not found: ' + fn };
    // google.script.run thật trả null cho CẢ phản hồi nếu bên trong có giá trị kiểu Date.
    const hasDate = (v) => v instanceof Date || (v && typeof v === 'object' && Object.values(v).some(hasDate));
    try { const r = mod[fn](...args); return r === undefined || hasDate(r) ? null : JSON.parse(JSON.stringify(r)); }
    catch (e) { return { __throw: e.message }; }
  });
  // Giả lập google.script.run (bất đồng bộ, chuỗi withSuccessHandler/withFailureHandler như GAS thật).
  // Chèn thẳng vào <head> vì addInitScript không chạy với page.setContent.
  const STUB = `<script>(function(){const mk=(s,f)=>new Proxy({},{get(_,p){if(p==='withSuccessHandler')return(h)=>mk(h,f);if(p==='withFailureHandler')return(h)=>mk(s,h);if(p==='withUserObject')return()=>mk(s,f);return(...a)=>{window.__gas(p,a).then((r)=>{if(r&&r.__throw){if(f)f(new Error(r.__throw));}else if(s)s(r);});};}});window.google={script:{run:mk(null,null)}};})();</script>`;
  const html = fs.readFileSync(path.join(__dirname, '..', 'Index.html'), 'utf8');
  await page.setContent(html.replace('<head>', '<head>' + STUB), { waitUntil: 'domcontentloaded' });
  await sleep(800);

  console.log('\n=== Dashboard ===');
  const dashText = await page.textContent('#dashboard');
  check('Dashboard tải được, có thẻ tổng hợp đồng', /Tổng số hợp đồng/.test(dashText));

  console.log('\n=== Hợp đồng: XSS + dấu nháy ===');
  await page.click('.tab[data-t="hopdong"]');
  await sleep(500);
  const hdRows = await page.$$eval('#hopDongTable tbody tr', (t) => t.length);
  check('Bảng hợp đồng hiển thị 3 dòng', hdRows === 3, 'thực tế ' + hdRows);
  const hdHtml = await page.innerHTML('#hopDongTable tbody');
  check('Tên chủ rừng chứa <img> được hiển thị dạng chữ (đã escape)', hdHtml.includes('&lt;img'));
  check('Không có thẻ <img> thật nào bị chèn vào bảng', (await page.$$eval('#hopDongTable img', (e) => e.length)) === 0);
  const opts = await page.$$eval('#selectHopDong option', (o) => o.map((x) => x.value));
  check('Dropdown giữ nguyên giá trị gốc HD"002 (escape không làm sai dữ liệu)', opts.includes(QUOTE_HD), JSON.stringify(opts));

  const idx = await page.$$eval('#hopDongTable tbody tr', (t, q) => t.findIndex((r) => r.textContent.includes(q)), QUOTE_HD);
  await page.click(`#hopDongTable tbody tr:nth-child(${idx + 1}) button`);
  await sleep(500);
  check('Nút "Chi tiết" (xemChiTietByIdx) gọi đúng getHoSoChiTiet("HD\\"002")', calls.some(([f, a]) => f === 'getHoSoChiTiet' && a[0] === QUOTE_HD));
  check('Popup chi tiết không thực thi payload GPS/ảnh', await page.evaluate(() => window.__xss === undefined && document.querySelectorAll('img[src="x"]').length === 0));
  await page.evaluate(() => { document.getElementById('chiTietModalBackdrop').style.display = 'none'; });

  console.log('\n=== Submit form Đánh giá hợp đồng ===');
  await page.selectOption('#selectHopDong', QUOTE_HD);
  await page.fill('#formDanhGia input[name="NguoiDanhGia"]', '=1+1');
  const disabledDuringSubmit = await page.evaluate(() => {
    const f = document.getElementById('formDanhGia');
    f.dispatchEvent(new Event('submit', { cancelable: true }));
    return f.querySelector('button[type="submit"]').disabled;
  });
  check('Nút Lưu bị khóa ngay khi đang gửi (chặn double-submit)', disabledDuringSubmit === true);
  await page.waitForSelector('#dgMsg .msg', { timeout: 3000 });
  check('Hiện thông báo thành công', (await page.$('#dgMsg .msg.ok')) !== null, await page.innerHTML('#dgMsg'));
  check('Nút Lưu được mở lại sau khi xong', await page.$eval('#formDanhGia button[type="submit"]', (b) => !b.disabled));
  const dgRows = env.ownSS.getSheetByName('DanhGiaHopDong').rows;
  const last = dgRows[dgRows.length - 1];
  check('Dữ liệu ghi đúng MaHopDong = HD"002', last[0] === QUOTE_HD, JSON.stringify(last));
  check('Giá trị "=1+1" được ghi dạng text (chặn công thức)', last.includes("'=1+1"), JSON.stringify(last));

  console.log('\n=== Lỗi phía server phải hiện ra, không báo thành công giả ===');
  await page.evaluate(() => { document.getElementById('dgMsg').innerHTML = ''; });
  const sheetBackup = env.ownSS.sheets.DanhGiaHopDong;
  delete env.ownSS.sheets.DanhGiaHopDong;
  await page.selectOption('#selectHopDong', 'HD003');
  await page.click('#formDanhGia button[type="submit"]');
  await page.waitForSelector('#dgMsg .msg', { timeout: 3000 });
  check('Mất sheet -> hiện thông báo LỖI (không phải "Đã lưu")', (await page.$('#dgMsg .msg.err')) !== null, await page.innerHTML('#dgMsg'));
  check('Nút Lưu được mở lại sau lỗi', await page.$eval('#formDanhGia button[type="submit"]', (b) => !b.disabled));
  env.ownSS.sheets.DanhGiaHopDong = sheetBackup;

  await page.evaluate(() => { document.getElementById('dgMsg').innerHTML = ''; document.getElementById('selectHopDong').value = ''; document.getElementById('formDanhGia').dispatchEvent(new Event('submit', { cancelable: true })); });
  await page.waitForSelector('#dgMsg .msg', { timeout: 3000 });
  const vmsg = await page.innerHTML('#dgMsg');
  check('Bỏ qua kiểm tra trình duyệt, gửi thiếu Số HĐ -> server từ chối "Thiếu trường bắt buộc"', /Thiếu trường bắt buộc/.test(vmsg), vmsg);

  console.log('\n=== Danh sách có cột ngày (Sheets lưu kiểu Date) ===');
  await page.click('.tab[data-t="thamdinh"]');
  await page.waitForTimeout(2500);
  const tdText = await page.textContent('#tdhdTable tbody');
  const tdErr = await page.$eval('#globalError', (e) => (getComputedStyle(e).display !== 'none' ? e.textContent : '')).catch(() => '');
  check('Tab Thẩm định tải được bản ghi có ngày, không báo "Phản hồi bất thường"', !tdErr && tdText.includes('HD003') && tdText.includes('04/10/2026'), (tdErr || tdText).slice(0, 200));
  await page.click('.tab[data-t="dashboard"]');
  await page.waitForTimeout(2500);
  const dbErr = await page.$eval('#globalError', (e) => (getComputedStyle(e).display !== 'none' ? e.textContent : '')).catch(() => '');
  check('Dashboard tải được khi Tiến độ có ngày thực tế', !dbErr, dbErr.slice(0, 200));

  console.log('\n=== Lô hàng mua vào: định dạng ngày + lọc theo ngày ===');
  await page.click('.tab[data-t="lohang"]');
  await page.waitForTimeout(800);
  const lhText = await page.textContent('#lohangTable tbody').catch(() => page.textContent('section#lohang'));
  check('Cột Ngày nhập hiện "01/08/2026", không có "00:00"', lhText.includes('01/08/2026') && !lhText.includes('00:00'), lhText.slice(0, 200));
  await page.fill('#fLoHangTu', '2026-08-02');
  await page.fill('#fLoHangDen', '2026-08-03');
  await page.evaluate(() => renderLoHang());
  const lhLoc = await page.textContent('section#lohang');
  check('Lọc 02/08–03/08 giữ đúng 2 ngày đó (không hiểu nhầm ngày/tháng)', lhLoc.includes('02/08/2026') && lhLoc.includes('03/08/2026') && !lhLoc.includes('01/08/2026') && !lhLoc.includes('04/08/2026'), lhLoc.slice(0, 300));
  await page.fill('#fLoHangTu', ''); await page.fill('#fLoHangDen', '');

  console.log('\n=== Duyệt toàn bộ tab ===');
  const tabs = await page.$$eval('.tab[data-t]', (t) => t.map((x) => x.dataset.t));
  for (const t of tabs) { await page.click(`.tab[data-t="${t}"]`); await sleep(250); }
  console.log('   đã mở ' + tabs.length + ' tab');
  check('Không có lỗi JavaScript nào khi mở tất cả tab', errors.length === 0, errors.join(' | '));
  check('Không payload XSS nào được thực thi trên toàn trang', await page.evaluate(() => window.__xss === undefined));
  check('Không có <img src=x> thật nào trên toàn trang', (await page.$$eval('img[src="x"]', (e) => e.length)) === 0);
  const notFound = calls.filter(([f]) => typeof mod[f] !== 'function').map(([f]) => f);
  check('Mọi hàm client gọi đều tồn tại trên server', notFound.length === 0, notFound.join(','));
  const ge = await page.$eval('#globalError', (e) => (getComputedStyle(e).display !== 'none' ? e.textContent : '')).catch(() => '');
  check('Không có thông báo lỗi chung (globalError) hiện ra', !ge, ge);

  await browser.close();
  console.log(`\n===== UI TEST: ${pass} PASS / ${fail} FAIL =====`);
  process.exit(fail ? 1 : 0);
})().catch((e) => { console.error(e); process.exit(2); });
