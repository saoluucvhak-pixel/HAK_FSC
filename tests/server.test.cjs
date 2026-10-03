'use strict';
// Test logic phía server (Code.gs) trên Google Sheets giả lập. Chạy: node tests/server.test.cjs
const { buildMockEnv } = require('./mock_gas.cjs');
const env = buildMockEnv();
const mod = require('./load_code.cjs')();
const { SS_IDS } = mod;

let pass = 0, fail = 0;
function check(name, cond, detail) {
  if (cond) { pass++; console.log(`✅ ${name}`); }
  else { fail++; console.log(`❌ ${name}${detail ? ' — ' + detail : ''}`); }
}
function section(t) { console.log('\n=== ' + t + ' ==='); }

function setupExternalSheets() {
  const hdmb = global.SpreadsheetApp.openById(SS_IDS.HDMB);
  hdmb._addSheet('HD_RUNG', [
    ['SoHopDong', 'MaRung', 'HoVaTenChuRung', 'ID_RUNG', 'SoCCCD'],
    ['HD001', 'R001', 'Nguyễn Văn A', 'GPS001', '201xxxxxxxx'],
    ['HD001', 'R002', 'Nguyễn Văn A', 'GPS002', '201xxxxxxxx'],
    ['HD002', 'R003', 'Trần Thị B', 'GPS003', '202xxxxxxxx'],
    ['HD003', 'R004', 'Lê Văn C', 'GPS004', '203xxxxxxxx'],
  ]);
  hdmb._addSheet('BaoCao_KiemTra', [
    ['Số HĐ', 'Kết quả', 'Hồ sơ còn thiếu', 'Cảnh báo'],
    ['HD001', 'Đầy đủ', '', ''],
    ['HD002', 'Thiếu CCCD', 'CCCD', 'Cần bổ sung'],
  ]);
  hdmb._addSheet('HD_Picture', [['ID_HD', 'Picture1'], ['R001', 'https://img/1.jpg']]);
  hdmb._addSheet('HD_GPS', [['ID_KEY_GPS', 'Latitude_Vĩ độ', 'Longtitude_Kinh độ', 'Địa điểm tra GPS'], ['GPS001', '15.9', '108.2', 'Quế Sơn']]);
  hdmb._addSheet('HD_STK', [['Số hợp đồng', 'Họ và tên người được ủy quyền', 'Số TK nhận tiền', 'Ngân hàng', 'Ủy quyền thanh toán']]);
  hdmb._addSheet('HD_NCC', [['SoHopDong'], ['HD001']]);

  const hosokeo = global.SpreadsheetApp.openById(SS_IDS.HOSOKEO);
  const header = ['STT', 'Ngày nhập', 'Mã hợp đồng', 'SỐ PHIẾU CÂN', 'Số BKLS', 'Họ và tên chủ rừng', 'ĐỊA CHỈ RỪNG', 'Khối lượng (Tấn)', 'Đơn giá', 'Thanh toán', 'Nguồn gốc'];
  const rows = [header];
  for (let i = 1; i <= 10; i++) {
    rows.push([i, new Date(2026, 7, i), 'HD00' + ((i % 3) + 1), 'PC' + i, 'BKLS' + i, 'Chủ rừng ' + i, 'Địa chỉ ' + i, 10 + i, 1500000, 'Đã TT', i % 2 === 0 ? 'PS' : 'DT']);
  }
  hosokeo._addSheet('HoSoKeo_DN', rows);
  hosokeo._addSheet('HoSoRung_DN', [['Mã hợp đồng'], ['HD001']]);
  hosokeo._addSheet('ToaDoRung_DN', [['Số hợp đồng'], ['HD001']]);

  const xuathang = global.SpreadsheetApp.openById(SS_IDS.XUATHANG);
  const xh = [['STT', 'Số phiếu', 'Ngày giờ cân 1', 'Số BKLS', 'Khối lượng (Tấn)', 'Khối lượng (M3)', 'Đơn vị vận chuyển', 'SỐ TKHQ', 'TÀU XUẤT']];
  for (let i = 1; i <= 5; i++) xh.push([i, 'XH' + i, new Date(2026, 7, i + 2), 'BKLS' + i, 8, 12, 'Vận tải ' + i, 'TK' + i, 'Tau ' + i]);
  xuathang._addSheet('NL_PC_XH', xh);
  xuathang._addSheet('NL_DH_XB', [['STT', 'Khách hàng'], [1, 'ABC']]);

  const phieuCan = global.SpreadsheetApp.openById(SS_IDS.PHIEUCAN);
  phieuCan._addSheet('PhieuCan_DN', [['STT'], [1]]);

  const khaosat = global.SpreadsheetApp.openById(SS_IDS.KHAOSAT_FORM);
  khaosat._addSheet('KhaoSat_FSC', [['Timestamp', 'Người khảo sát'], ['2026-08-01', 'A']]);
}
setupExternalSheets();

section('1) initOwnSheets / checkKetNoi');
{
  const chk0 = mod.checkKetNoi();
  check('checkKetNoi báo ownSheetsReady=false TRƯỚC khi initOwnSheets', chk0.ownSheetsReady === false);
  const initMsg = mod.initOwnSheets();
  console.log('   initOwnSheets() =>', initMsg);
  const chk1 = mod.checkKetNoi();
  check('checkKetNoi báo ownSheetsReady=true SAU khi initOwnSheets', chk1.ownSheetsReady === true);
  check('checkKetNoi báo tất cả ' + chk1.external.length + ' sheet nguồn ok=true (không còn TongHop_HopDong)', chk1.external.every((r) => r.ok === true), JSON.stringify(chk1.external.filter((r) => !r.ok)));
}

section('2) getHopDongList — ghép & khử trùng + KHÔNG còn field ChenhLech');
{
  const r = mod.getHopDongList();
  check('getHopDongList chạy OK', r.ok, r.error);
  check('Khử trùng đúng -> còn 3 hợp đồng', r.data.length === 3, 'thực tế = ' + r.data.length);
  const hd001 = r.data.find((x) => x.SoHopDong === 'HD001');
  check('Đã bỏ hẳn field ChenhLechDienTichPhanTram', hd001 && !('ChenhLechDienTichPhanTram' in hd001));
}

section('3) getDashboardData — xác nhận hành vi đếm đã sửa + đã bọc _safe cho add/update');
{
  const r1 = mod.addDanhGiaHopDong({ MaHopDong: 'HD001', KetLuanGiamSat: 'ĐẠT' });
  check('addDanhGiaHopDong giờ trả {ok:true,...} (đã bọc _safe)', r1 && r1.ok === true, JSON.stringify(r1));
  mod.addDanhGiaHopDong({ MaHopDong: 'HD001', KetLuanGiamSat: 'ĐẠT CÓ ĐIỀU KIỆN' });
  mod.addDanhGiaHopDong({ MaHopDong: 'HD002', KetLuanGiamSat: 'KHÔNG ĐẠT' });

  const r2 = mod.addGiamSat({ KyGiamSat: 'Q3-2026' });
  check('addGiamSat giờ trả {ok:true,...} (đã bọc _safe)', r2 && r2.ok === true, JSON.stringify(r2));

  const r3 = mod.addDanhGiaRuiRoRung({ MaRung: 'R001', SoHopDong: 'HD001', RuiRoCNRA: 'Không đáng kể' });
  check('addDanhGiaRuiRoRung giờ trả {ok:true,...} (đã bọc _safe)', r3 && r3.ok === true, JSON.stringify(r3));

  const r4 = mod.updateTienDo('0. Chuẩn bị', '2026-09-01', 'Đúng tiến độ', 'test');
  check('updateTienDo giờ trả {ok:true,...} (đã bọc _safe)', r4 && r4.ok === true, JSON.stringify(r4));
  const r5 = mod.updateTienDo('Giai đoạn không tồn tại', '2026-09-01', 'x', 'y');
  check('updateTienDo với giai đoạn không tồn tại trả {ok:false,...} thay vì throw ra ngoài hoặc trả false im lặng', r5 && r5.ok === false, JSON.stringify(r5));

  const dash = mod.getDashboardData();
  const d = dash.data;
  console.log('   Kết quả Dashboard:', JSON.stringify({ tongHD: d.tongHD, dat: d.dat, datCoDK: d.datCoDK, khongDat: d.khongDat, chuaDanhGia: d.chuaDanhGia }));
  check('chuaDanhGia=1 đúng (HD003 chưa đánh giá)', d.chuaDanhGia === 1);
  check('HD001 chỉ tính kết luận GẦN NHẤT (ĐẠT CÓ ĐIỀU KIỆN) — dat=0,datCoDK=1,khongDat=1', d.dat === 0 && d.datCoDK === 1 && d.khongDat === 1);
  check('Tổng dat+datCoDK+khongDat+chuaDanhGia = tongHD', (d.dat + d.datCoDK + d.khongDat + d.chuaDanhGia) === d.tongHD);
  check('Đã bỏ hẳn field chenhLechCanhBao', !('chenhLechCanhBao' in d));
}

section('4) Mô phỏng lỗi server để kiểm tra client KHÔNG còn báo "thành công" giả (C1)');
{
  // Xóa tạm sheet DanhGiaHopDong để giả lập lỗi "mất sheet" (bị đổi tên/xóa)
  const backup = env.ownSS.sheets['DanhGiaHopDong'];
  delete env.ownSS.sheets['DanhGiaHopDong'];
  const rFail = mod.addDanhGiaHopDong({ MaHopDong: 'HD999', KetLuanGiamSat: 'ĐẠT' });
  check('Khi sheet bị mất, addDanhGiaHopDong trả {ok:false, error:...} rõ ràng (không throw ra ngoài, không trả true giả)', rFail && rFail.ok === false, JSON.stringify(rFail));
  env.ownSS.sheets['DanhGiaHopDong'] = backup;
}

section('5) getMauWordLinks, getTonKhoBaoCao vẫn hoạt động bình thường (hồi quy)');
{
  const r = mod.getMauWordLinks();
  check('getMauWordLinks chạy OK', r.ok && r.data.length === 2, JSON.stringify(r));
  const r2 = mod.getTonKhoBaoCao(null, null);
  check('getTonKhoBaoCao chạy OK', r2.ok, r2.error);
}

section('6) PHASE 2 — H4: validate khóa nghiệp vụ bắt buộc (server-side)');
{
  const r1 = mod.addKhaoSat({ NguoiKhaoSat: 'A' }); // thiếu SoHopDong
  check('addKhaoSat thiếu SoHopDong -> {ok:false} với thông báo rõ ràng', r1 && r1.ok === false && /SoHopDong/.test(r1.error), JSON.stringify(r1));
  const r2 = mod.addDanhGiaRuiRoRung({ MaRung: 'R001' }); // thiếu SoHopDong
  check('addDanhGiaRuiRoRung thiếu SoHopDong -> {ok:false}', r2 && r2.ok === false && /SoHopDong/.test(r2.error), JSON.stringify(r2));
  const r3 = mod.addTieuChiThamDinh({ SoHopDong: 'HD001', NguoiThamDinh: 'B' });
  check('addTieuChiThamDinh đủ SoHopDong -> {ok:true}', r3 && r3.ok === true, JSON.stringify(r3));
}

section('7) PHASE 2 — H5: chặn formula injection khi ghi vào Sheet');
{
  const payload = { SoHopDong: 'HD001', NguoiKhaoSat: '=HYPERLINK("http://evil")' };
  const r = mod.addKhaoSat(payload);
  check('addKhaoSat với giá trị bắt đầu bằng "=" vẫn lưu thành công', r && r.ok === true, JSON.stringify(r));
  const sh = mod._own('KhaoSatThamVan');
  const rows = sh.getDataRange().getValues();
  const last = rows[rows.length - 1];
  const headers = rows[0];
  const idx = headers.indexOf('NguoiKhaoSat');
  check('Giá trị đã được thêm dấu nháy đơn để ép kiểu text, không còn bắt đầu bằng "="', typeof last[idx] === 'string' && last[idx][0] === "'" , JSON.stringify(last[idx]));
}

section('8) PHASE 2 — H2: các hàm sinh mã tự động vẫn hoạt động đúng sau khi thêm LockService');
{
  const r1 = mod.addRuiRoTrienKhai({ MoTa: 'Rủi ro test' });
  check('addRuiRoTrienKhai vẫn trả {ok:true} và tự sinh MaRuiRo', r1 && r1.ok === true, JSON.stringify(r1));
  const r2 = mod.addLenhDieuDong({ LoaiHang: 'Gỗ' });
  check('addLenhDieuDong vẫn trả {ok:true} và tự sinh SoLenh', r2 && r2.ok === true, JSON.stringify(r2));
  const r3 = mod.addPhieuKhacPhuc({ KhuVucBoPhan: 'Kho' });
  check('addPhieuKhacPhuc vẫn trả {ok:true} và tự sinh SoPhieu', r3 && r3.ok === true, JSON.stringify(r3));
}

section('9) PHASE 2 — H3: _extReadLastN không bỏ sót dòng khi cột A (STT) trống ở dòng cuối');
{
  const hosokeo = global.SpreadsheetApp.openById(mod.SS_IDS.HOSOKEO);
  const existing = hosokeo.getSheetByName('HoSoKeo_DN');
  // Thêm 1 dòng mới có STT (cột A) TRỐNG nhưng các cột khác có dữ liệu thật — mô phỏng lỗi nhập liệu
  existing.appendRow(['', new Date(2026, 7, 20), 'HD003', 'PC_MOI', 'BKLS_MOI', 'Chủ rừng mới', 'Địa chỉ mới', 99, 1500000, 'Đã TT', 'DT']);
  const rows = mod._extReadLastN('HOSOKEO', 'HoSoKeo_DN', 3);
  const found = rows.some((r) => r['Số BKLS'] === 'BKLS_MOI');
  check('Dòng mới (STT trống) vẫn được đọc ra dù cột A trống', found, JSON.stringify(rows));
}

section('10) PHASE 3 — CacheService: giảm đọc chéo sheet ngoài + tự xóa cache khi có thay đổi');
{
  Object.keys(env.cacheStore).forEach((k) => delete env.cacheStore[k]);

  const d1 = mod.getDashboardData();
  check('getDashboardData lần đầu chạy OK và ghi vào cache', d1.ok && ('dashboardData' in env.cacheStore));
  const d2 = mod.getDashboardData();
  check('getDashboardData lần 2 trả đúng dữ liệu đã cache', d2.ok && JSON.stringify(d2.data) === JSON.stringify(d1.data));

  const hd1 = mod.getHopDongList();
  check('getHopDongList ghi vào cache', hd1.ok && ('hopDongList' in env.cacheStore));
  const rl1 = mod.getRungList();
  check('getRungList ghi vào cache', rl1.ok && ('rungList' in env.cacheStore));
  const tt1 = mod.getTinhTrangKhaoSatRung();
  check('getTinhTrangKhaoSatRung ghi vào cache', tt1.ok && ('tinhTrangKhaoSat' in env.cacheStore));
  const tk1 = mod.getTonKhoBaoCao(null, null);
  check('getTonKhoBaoCao ghi vào cache theo đúng khóa tham số', tk1.ok && ('tonKho__' in env.cacheStore));

  const addRes = mod.addKhaoSat({ SoHopDong: 'HD002', NguoiKhaoSat: 'Tester' });
  check('addKhaoSat lưu thành công', addRes.ok, JSON.stringify(addRes));
  check('Sau khi addKhaoSat, cache "dashboardData" đã tự xóa (không báo dữ liệu cũ)', !('dashboardData' in env.cacheStore));
  check('Sau khi addKhaoSat, cache "tinhTrangKhaoSat" đã tự xóa', !('tinhTrangKhaoSat' in env.cacheStore));

  const d3 = mod.getDashboardData();
  check('getDashboardData sau khi cache bị xóa tính lại đúng (chuaKhaoSat giảm)', d3.ok && d3.data.chuaKhaoSat < d1.data.chuaKhaoSat, JSON.stringify({ truoc: d1.data.chuaKhaoSat, sau: d3.data.chuaKhaoSat }));
}

console.log(`\n===== TỔNG KẾT: ${pass} PASS / ${fail} FAIL =====`);
process.exit(fail > 0 ? 1 : 0);
