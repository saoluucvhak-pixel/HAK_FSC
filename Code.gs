/**
 * HAK GROUP — HỆ THỐNG KIỂM SOÁT QUY TRÌNH FSC (QT-FSC-01 / KH-FSC-01) — v2
 * Backend Google Apps Script — dán toàn bộ file này vào Extensions > Apps Script > Code.gs
 *
 * KHÁC BIỆT SO VỚI v1: hệ thống này KHÔNG lưu trùng dữ liệu — nó ĐỌC trực tiếp
 * 5 Google Sheet nghiệp vụ đang dùng thật (hợp đồng, phiếu cân, hồ sơ keo, xuất
 * hàng) và chỉ lưu MỚI phần chưa từng có: kết luận giám sát định kỳ, đánh giá
 * rủi ro CNRA theo rừng, và theo dõi tiến độ triển khai KH-FSC-01.
 *
 * Chạy hàm initOwnSheets() MỘT LẦN trước khi dùng.
 */

// ============ CẤU HÌNH 5 SPREADSHEET NGUỒN (chỉ đọc) ============
const SS_IDS = {
  HDMB: '1cv11ORWuAF3Sit4f-kA0xrP6-ab4SF-7LEdkCvGi_gI',      // Hợp đồng, hồ sơ rừng, GPS, ảnh, checklist có sẵn
  PHIEUCAN: '1vqMVxccBA7zlAMHrGsVBydGFwZJ6QuDZW10zJ74V29g',  // Phiếu cân nhà máy Đà Nẵng
  DNTT: '1oUm87_gbDbnuPc_We0dyZ_e4kHXBHXs95AQAxp5okYo',      // Đề nghị thanh toán, sổ nhập tổng hợp, đối soát công nợ
  HOSOKEO: '1PfXmgnO4ad1Aourjcjoh7hZxL6mVuywL73wHle0oq6I',   // Hồ sơ keo (mua vào), hồ sơ rừng, tọa độ rừng
  XUATHANG: '1ZZ2iUwkkKe8wXdztA7mL-v9j6fmgY5c5rlDdI1sNoAk',  // Đơn hàng & xuất hàng
  KHAOSAT_FORM: '1_rAkzrUDbuTM3bSjT1V04bTQ5wjKeWK4veT8AeWO3Bw', // Google Form phản hồi Báo cáo khảo sát thực địa
};

// ============ THƯ MỤC GOOGLE DRIVE CHỨA MẪU WORD GỐC (chỉ đọc) ============
const MAU_WORD_FOLDER_ID = '18YF0RedMZhozPEk2U0AUvYJCBwpOtbtY';

/** Liệt kê toàn bộ file trong thư mục mẫu Word trên Drive, kèm link xem trực tiếp.
 * Dùng DriveApp (không phải SpreadsheetApp) vì đây là thư mục Drive thường, không phải Sheet. */
function getMauWordLinks() {
  return _safe(() => {
    const folder = DriveApp.getFolderById(MAU_WORD_FOLDER_ID);
    const files = folder.getFiles();
    const result = [];
    while (files.hasNext()) {
      const f = files.next();
      result.push({ name: f.getName(), url: f.getUrl(), id: f.getId() });
    }
    result.sort((a, b) => a.name.localeCompare(b.name, 'vi'));
    return result;
  });
}

// ============ SHEET TỰ TẠO TRONG FILE NÀY (dữ liệu MỚI, không trùng nguồn) ============
const OWN_HEADERS = {
  DanhGiaHopDong: ['MaHopDong', 'NgayDanhGia', 'KetLuanGiamSat', 'PhatHien', 'HanhDongKhacPhuc', 'NguoiDanhGia', 'GhiChu'],
  DanhGiaRuiRoRung: ['MaRung', 'SoHopDong', 'RuiRoCNRA', 'BienPhapRMM', 'NgayDanhGia', 'NguoiDanhGia', 'GhiChu'],
  GiamSatDinhKy: ['MaBaoCao', 'KyGiamSat', 'PhamViKiemTra', 'MucTieu1_DayDuHoSo', 'MucTieu2_HopPhap', 'MucTieu3_NhatQuan', 'MucTieu4_DungHan', 'MucTieu5_RuiRoFSC', 'PhatHienChinh', 'KetLuan', 'HanhDongKhacPhuc', 'NguoiPhuTrach', 'HanHoanThanh', 'NgayTaiKiemTra'],
  TienDoTrienKhai: ['GiaiDoan', 'MocKeHoach', 'NgayThucTeHoanThanh', 'ChenhLechTuan', 'DanhGia', 'GhiChu'],
  RuiRoTrienKhai: ['MaRuiRo', 'MoTa', 'MucDo', 'GiaiDoanLienQuan', 'BienPhapGiamThieu', 'NguoiPhuTrach', 'HanXuLy', 'TrangThai'],
  KhaoSatThamVan: ['SoHopDong', 'NgayKhaoSat', 'NguoiKhaoSat', 'CoBaoCaoKhaoSat', 'CoBBGiamSatTrongKhaiThac', 'CoBBGiamSatSauKhaiThac', 'ThamVanUBNDXa', 'ThamVanKiemLam', 'ThamVanChuRungLanCan', 'ThamVanNguoiDanDiaPhuong', 'ThamVanCongNhanKhaiThac', 'ThamVanNguoiVanChuyen', 'PhatHienRuiRo', 'GhiChu'],
  BaoCaoKhaoSat: ['SoHopDong', 'NgayKhaoSat', 'NguoiKhaoSat', 'DiaChiRungTrong', 'ToaDo', 'PhapLyThuaDat', 'MucDichSuDungDat', 'KhuVucRungTrong', 'DatDai', 'MoiTruong', 'LaoDong', 'CongDong', 'DuongVanChuyen', 'HoatDongVanChuyenGo', 'HoatDongKhaiThac', 'GiaTriBaoTon', 'XuLyThucBi', 'AnhKhaoSat', 'GhiChu'],
  BienBanGiamSatKhaiThac: ['SoHopDong', 'GiaiDoanGiamSat', 'NguoiKiemTra', 'NgayKiemTra', 'HoatDongKhaiThac', 'HoatDongVanChuyen', 'LaoDong', 'MoiTruong', 'GiaTriBaoTon', 'CacHoatDongKhac', 'PhuongAnPhongNgua', 'GhiChu'],
  ThamVanBenLienQuan: ['SoHopDong', 'LoaiBenLienQuan', 'TenNguoiDuocThamVan', 'ChucVu', 'NgayThamVan', 'NoiDungThamVan', 'KetLuan', 'GhiChu'],
  TieuChiThamDinhHopDong: ['SoHopDong', 'NgayThamDinh', 'MucDichSuDungDatHopLe', 'SoDoConHan', 'CCCDConHieuLuc', 'KhongPhaiRungTuNhienSau2020', 'KhongTranhChap', 'KhoiLuongKhongVuotSoDo', 'ThongBaoGiaDayDu', 'KetLuanThamDinh', 'NguoiThamDinh', 'GhiChu'],
  LenhDieuDongVanChuyen: ['SoLenh', 'NgayLenh', 'LoaiHang', 'DiemDi', 'DiemDen', 'BienKiemSoat', 'SoPhieuCan', 'SoBKLSGoc', 'NgayHoanThanhVanChuyen', 'NgayLapPhieuXuatKho', 'NguoiLapLenh', 'GhiChu'],
  TieuHaoCheBien: ['KyTheoDoi', 'SoBKLSXuatCheBien', 'KhoiLuongDauVaoTan', 'DinhMucTieuHaoPhanTram', 'DamThanhPhamTan', 'TieuHaoThucTePhanTram', 'GiaiTrinhChenhLech', 'NguoiGhiSo'],
  SoTheoDoiBaoCaoNhaNuoc: ['LoaiBaoCao', 'Ky', 'HanNop', 'NgayNopThucTe', 'DungHan', 'DauMoiTiepNhan', 'NguoiPhuTrach', 'GhiChu'],
  NghiemThuGiaiDoanTrienKhai: ['GiaiDoan', 'ThoiGianThucHien', 'DauRaKyVong', 'BangChungThucNhan', 'KhoangTrongBangChung', 'KetLuanNghiemThu', 'DieuKienChuyenGiaiDoan', 'NguoiNghiemThu', 'NgayNghiemThu'],
  KeHoachDanhGiaNoiBo: ['SoKeHoach', 'TuNgay', 'DenNgay', 'MucDichDanhGia', 'TongSoNhaCungCap', 'SoMauLayTheoCongThuc', 'TruongDoan', 'ThanhVienDoan', 'NoiDungDanhGia', 'NgayLap'],
  PhieuYeuCauKhacPhuc: ['SoPhieu', 'NgayDanhGia', 'NgayPhatHanh', 'KhuVucBoPhan', 'NoiDungKhongPhuHop', 'PhanTichNguyenNhan', 'NguoiPhanTich', 'HanhDongKhacPhuc', 'DuKienNgayKetThuc', 'TrangThai', 'NgayDongThucTe', 'GhiChu'],
  BienBanXemXetHeThongDDS: ['NgayXemXet', 'DiaDiem', 'ThanhPhanThamDu', 'KetQuaDanhGiaNoiBoNam', 'KetLuanThayDoi', 'NoiDungTruocThayDoi', 'NoiDungSauThayDoi', 'NguoiLap'],
  PheDuyetNangLucDanhGiaVien: ['HoTen', 'ChucVu', 'TrinhDoChuyenNganh', 'DatYeuCauTrinhDo', 'DatYeuCauKinhNghiem', 'DatYeuCauPhamChat', 'DatYeuCauNangLuc', 'KetLuanPheDuyet', 'NgayPheDuyet', 'NguoiPheDuyet'],
};

function initOwnSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(OWN_HEADERS).forEach((name) => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, OWN_HEADERS[name].length).setValues([OWN_HEADERS[name]]);
      sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, OWN_HEADERS[name].length).setFontWeight('bold').setBackground('#D9E2F3');
    }
  });
  const tienDo = ss.getSheetByName('TienDoTrienKhai');
  if (tienDo.getLastRow() === 1) {
    tienDo.getRange(2, 1, 5, 2).setValues([
      ['0. Chuẩn bị', 'Tuần 1-2'],
      ['1. Đào tạo & Thí điểm', 'Tuần 3-6'],
      ['2. Triển khai toàn diện', 'Tuần 7-12'],
      ['3. Đánh giá nội bộ', 'Tuần 13-16'],
      ['4. Vận hành ổn định', 'Từ tuần 17'],
    ]);
  }
  return 'OK — đã khởi tạo ' + Object.keys(OWN_HEADERS).length + ' sheet dữ liệu mới.';
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('HAK Group — Kiểm soát FSC')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ============ TIỆN ÍCH ĐỌC SHEET NGUỒN (bên ngoài, chỉ đọc) ============
function _ext(ssKey, sheetName) {
  return SpreadsheetApp.openById(SS_IDS[ssKey]).getSheetByName(sheetName);
}
function _toObjRows(values) {
  if (values.length < 2) return [];
  const headers = values[0];
  return values.slice(1)
    .filter((r) => r.some((c) => c !== '' && c !== null))
    .map((r) => { const o = {}; headers.forEach((h, i) => { o[h] = r[i]; }); return o; });
}
/** Đọc toàn bộ sheet — chỉ dùng cho sheet nhỏ (dưới ~2000 dòng), ví dụ BaoCao_KiemTra, TongHop_HopDong, HD_RUNG */
function _extReadAll(ssKey, sheetName) {
  const sh = _ext(ssKey, sheetName);
  if (!sh) return [];
  return _toObjRows(sh.getDataRange().getValues());
}
/** Đọc N dòng GẦN NHẤT — TỐI ƯU cho sổ lớn (HoSoKeo_DN, PhieuCan_DN, NL_PC_XH hàng nghìn dòng).
 * Thay vì đọc hết cả sheet (chậm, đặc biệt vì đây là spreadsheet KHÁC — cross-spreadsheet access
 * chậm hơn nhiều so với đọc trong cùng file), ta chỉ quét CỘT ĐẦU TIÊN (thường là STT/Số thứ tự,
 * luôn có giá trị ở mọi dòng thật) để tìm dòng cuối cùng có dữ liệu thật, rồi mới đọc đúng N dòng
 * cuối trên TOÀN BỘ cột — nhanh hơn nhiều vì vùng đọc nhỏ. */
function _extReadLastN(ssKey, sheetName, n) {
  const sh = _ext(ssKey, sheetName);
  if (!sh) return [];
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2) return [];

  // Bước 1: quét CHỈ cột A (nhanh) để tìm dòng cuối THẬT có dữ liệu
  const colA = sh.getRange(2, 1, lastRow - 1, 1).getValues();
  let trueLastRow = 1;
  for (let i = colA.length - 1; i >= 0; i--) {
    if (colA[i][0] !== '' && colA[i][0] !== null) { trueLastRow = i + 2; break; }
  }
  if (trueLastRow <= 1) return []; // cột A trống hết — không có dữ liệu thật

  // Bước 2: chỉ đọc đúng N dòng cuối (vùng nhỏ, nhanh) + dòng header
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  const startRow = Math.max(2, trueLastRow - n + 1);
  const values = sh.getRange(startRow, 1, trueLastRow - startRow + 1, lastCol).getValues();
  return _toObjRows([headers, ...values]).reverse();
}

// ============ TIỆN ÍCH SHEET RIÊNG (đọc/ghi) ============
function _own(name) { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name); }
function _ownReadAll(name) { return _toObjRows(_own(name).getDataRange().getValues()); }
function _ownAppend(name, obj) {
  const sh = _own(name);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  sh.appendRow(headers.map((h) => (obj[h] !== undefined ? obj[h] : '')));
  return true;
}

// ============ BỌC LỖI: mọi hàm gọi từ client đi qua đây để lỗi hiện rõ thay vì im lặng ============
function _safe(fn) {
  try {
    return { ok: true, data: fn() };
  } catch (e) {
    return { ok: false, error: e.message + (e.stack ? ('\n' + e.stack) : '') };
  }
}

/** Kiểm tra nhanh: có mở được cả 5 file nguồn + đọc được sheet chính không.
 * Gọi hàm này đầu tiên khi nghi ngờ lỗi quyền truy cập. */
function checkKetNoi() {
  const targets = [
    ['HDMB', 'HD_NCC'], ['HDMB', 'BaoCao_KiemTra'], ['HDMB', 'TongHop_HopDong'], ['HDMB', 'HD_RUNG'],
    ['PHIEUCAN', 'PhieuCan_DN'],
    ['HOSOKEO', 'HoSoKeo_DN'], ['HOSOKEO', 'HoSoRung_DN'], ['HOSOKEO', 'ToaDoRung_DN'],
    ['XUATHANG', 'NL_PC_XH'],
    ['KHAOSAT_FORM', 'KhaoSat_FSC'],
  ];
  // Ghi chú: DNTT/DN không nằm trong danh sách kiểm tra vì hiện dashboard KHÔNG đọc sheet này
  // (đã thay bằng HD_RUNG + HoSoKeo_DN làm bảng gốc). Nếu sau này cần đối chiếu công nợ từ
  // DNTT_GK_DN, sẽ bổ sung lại đúng tên sheet thật tại thời điểm đó.
  const results = targets.map(([ssKey, sheetName]) => {
    try {
      const ss = SpreadsheetApp.openById(SS_IDS[ssKey]);
      const sh = ss.getSheetByName(sheetName);
      if (!sh) return { ssKey, sheetName, ok: false, error: 'Mở được file nhưng KHÔNG thấy sheet tên "' + sheetName + '" — kiểm tra lại tên sheet có đúng chính tả/dấu cách không.' };
      return { ssKey, sheetName, ok: true, soDong: sh.getLastRow() };
    } catch (e) {
      return { ssKey, sheetName, ok: false, error: e.message };
    }
  });
  let ownOk = true, ownError = '';
  try {
    Object.keys(OWN_HEADERS).forEach((name) => {
      if (!_own(name)) throw new Error('Chưa có sheet "' + name + '" trong file này — cần chạy initOwnSheets() trước.');
    });
  } catch (e) { ownOk = false; ownError = e.message; }
  return { external: results, ownSheetsReady: ownOk, ownError };
}

/** Hàm đơn giản nhất có thể — không đụng đến Sheet nào cả. Dùng để test xem
 * kênh gọi google.script.run từ trình duyệt có hoạt động hay không. */
function pingTest() {
  return 'PONG lúc ' + new Date().toLocaleString();
}

// ============ DEBUG: chạy trực tiếp bằng nút ▶ Run trong Apps Script, xem ở View > Logs ============
function debugHoSoKeo() {
  const sh = _ext('HOSOKEO', 'HoSoKeo_DN');
  Logger.log('Sheet tìm thấy: ' + (sh ? sh.getName() : 'KHÔNG TÌM THẤY'));
  if (!sh) return;
  Logger.log('getLastRow = ' + sh.getLastRow() + ' | getLastColumn = ' + sh.getLastColumn());
  const values = sh.getDataRange().getValues();
  Logger.log('Tổng số dòng đọc được (kể cả header) = ' + values.length);
  Logger.log('Dòng header: ' + JSON.stringify(values[0]));
  Logger.log('Dòng dữ liệu thứ 2 (mẫu): ' + JSON.stringify(values[1]));
  Logger.log('Dòng cuối cùng: ' + JSON.stringify(values[values.length - 1]));
  const rows = _toObjRows(values);
  Logger.log('Số dòng sau khi lọc rỗng = ' + rows.length);
  Logger.log('3 dòng cuối sau lọc: ' + JSON.stringify(rows.slice(-3)));
}
function debugXuatHang() {
  const sh = _ext('XUATHANG', 'NL_PC_XH');
  Logger.log('Sheet tìm thấy: ' + (sh ? sh.getName() : 'KHÔNG TÌM THẤY'));
  if (!sh) return;
  Logger.log('getLastRow = ' + sh.getLastRow() + ' | getLastColumn = ' + sh.getLastColumn());
  const values = sh.getDataRange().getValues();
  Logger.log('Tổng số dòng đọc được (kể cả header) = ' + values.length);
  Logger.log('Dòng header: ' + JSON.stringify(values[0]));
  Logger.log('Dòng dữ liệu thứ 2 (mẫu): ' + JSON.stringify(values[1]));
  const rows = _toObjRows(values);
  Logger.log('Số dòng sau khi lọc rỗng = ' + rows.length);
}
/** Giả lập ĐÚNG những gì webapp gọi — nếu hàm này trả dữ liệu ở đây mà webapp vẫn trống,
 * chắc chắn 100% là do CHƯA DEPLOY PHIÊN BẢN MỚI, không phải do code sai. */
function debugGoiNhuWebapp() {
  const res = getLoHangGanDay(150);
  Logger.log('ok = ' + res.ok);
  if (!res.ok) { Logger.log('LỖI: ' + res.error); return; }
  Logger.log('Số dòng trả về = ' + res.data.length);
  Logger.log('Dòng đầu tiên: ' + JSON.stringify(res.data[0]));
}

/** Danh sách rừng (không gộp trùng — mỗi lô rừng riêng) để đổ vào dropdown chọn khi đánh giá CNRA */
function getRungList() {
  return _safe(() => {
    const rung = _extReadAll('HDMB', 'HD_RUNG');
    return rung.map((r) => ({
      MaRung: _safeStr(r.MaRung),
      SoHopDong: _safeStr(r.SoHopDong),
      ChuRung: _safeStr(r.HoVaTenChuRung),
    })).filter((r) => r.MaRung || r.SoHopDong);
  });
}

// ============ API: DANH SÁCH HỢP ĐỒNG / HỒ SƠ RỪNG (ghép từ dữ liệu có sẵn) ============
/**
 * Ghép BaoCao_KiemTra (kết quả kiểm tra hồ sơ có sẵn) + TongHop_HopDong (chênh lệch
 * diện tích ký vs GPS) + DanhGiaHopDong (kết luận giám sát FSC do mình đánh giá thêm),
 * theo khóa "Số HĐ". Đây là danh sách để Bước 1/2 (QT-FSC-01) tham chiếu.
 */
function getHopDongList() {
  return _safe(() => {
    const rung = _extReadAll('HDMB', 'HD_RUNG');                 // SoHopDong, HoVaTenChuRung, ... — bảng gốc, luôn có dữ liệu
    const kiemTra = _extReadAll('HDMB', 'BaoCao_KiemTra');       // Số HĐ, Kết quả, Hồ sơ còn thiếu, Cảnh báo — có thể rỗng nếu chưa chạy
    const tongHop = _extReadAll('HDMB', 'TongHop_HopDong');      // Số HĐ, Chênh lệch (%)
    const danhGia = _ownReadAll('DanhGiaHopDong');                // MaHopDong, KetLuanGiamSat, ...

    const kiemTraMap = {};
    kiemTra.forEach((r) => { kiemTraMap[String(r['Số HĐ'])] = r; });
    const tongHopMap = {};
    tongHop.forEach((r) => { tongHopMap[String(r['Số HĐ'])] = r; });
    const danhGiaMap = {};
    danhGia.forEach((r) => { danhGiaMap[String(r.MaHopDong)] = r; });

    // Loại trùng: HD_RUNG có thể có nhiều dòng cho cùng 1 SoHopDong (nhiều lô rừng/HĐ) — gộp theo SoHopDong
    const seen = {};
    const result = [];
    rung.forEach((r) => {
      const soHD = String(r.SoHopDong);
      if (seen[soHD]) return;
      seen[soHD] = true;
      const kt = kiemTraMap[soHD] || {};
      const th = tongHopMap[soHD] || {};
      const dg = danhGiaMap[soHD] || {};
      result.push({
        SoHopDong: soHD,
        MaRung: r.MaRung || (kt['Mã Rừng'] || ''),
        ChuRung: r.HoVaTenChuRung || kt['Chủ rừng'] || '',
        KetQuaHoSo: kt['Kết quả'] || '(chưa có dữ liệu kiểm tra)',
        HoSoConThieu: kt['Hồ sơ còn thiếu'] || '',
        CanhBao: kt['Cảnh báo'] || '',
        ChenhLechDienTichPhanTram: th['Chênh lệch (%)'] != null ? th['Chênh lệch (%)'] : '',
        KetLuanGiamSat: dg.KetLuanGiamSat || '',
        NgayDanhGia: dg.NgayDanhGia || '',
      });
    });
    return result;
  });
}

/** Chi tiết hồ sơ đính kèm cho 1 hợp đồng. Chuỗi khóa nối đúng theo thực tế dữ liệu:
 * Số hợp đồng → (một hợp đồng có thể có NHIỀU rừng) tất cả dòng HD_RUNG cùng SoHopDong,
 * mỗi dòng có 1 MaRung → MaRung đó khớp với ID_HD trong HD_Picture để lấy ảnh
 * (mỗi rừng có thể có nhiều ảnh, cột Picture1..Picture10). Riêng tọa độ GPS nối theo
 * ID_RUNG (HD_RUNG) ↔ ID_KEY_GPS (HD_GPS). HD_STK nối theo Số hợp đồng. */
function getHoSoChiTiet(soHopDong) {
  return _safe(() => {
    const rungRows = _extReadAll('HDMB', 'HD_RUNG').filter((r) => String(r.SoHopDong) === String(soHopDong));
    const allPictures = _extReadAll('HDMB', 'HD_Picture');
    const allGps = _extReadAll('HDMB', 'HD_GPS');
    const picCols = ['Picture1', 'Picture2', 'Picture3', 'Picture4', 'Picture5', 'Picture6', 'Picture7', 'Picture8', 'Picture9', 'Picture10'];

    const rungList = rungRows.map((rung) => {
      const anhRows = allPictures.filter((p) => String(p.ID_HD) === String(rung.MaRung));
      const anhFiles = [];
      anhRows.forEach((p) => picCols.forEach((c) => { if (p[c]) anhFiles.push(_safeStr(p[c])); }));
      const gps = allGps.find((g) => String(g.ID_KEY_GPS) === String(rung.ID_RUNG));
      return {
        MaRung: _safeStr(rung.MaRung),
        HoVaTenChuRung: _safeStr(rung.HoVaTenChuRung),
        SoCCCD: _safeStr(rung.SoCCCD),
        ThuongTru: _safeStr(rung.ThuongTru),
        DiaChiRungTrong: _safeStr(rung.DiaChiRungTrong),
        DienTich_m2: _safeStr(rung.DienTich_m2),
        HoSoNguonGoc: _safeStr(rung.HoSoNguonGoc),
        SoGiayTo: _safeStr(rung.SoGiayTo),
        NgayGiayTo: _safeStr(rung.NgayGiayTo),
        DinhKemGiayTo: _safeStr(rung.DinhKemGiayTo),
        SoAnh: anhFiles.length,
        AnhFiles: anhFiles,
        Gps: gps ? { Lat: _safeStr(gps['Latitude_Vĩ độ']), Lng: _safeStr(gps['Longtitude_Kinh độ']), DiaDiem: _safeStr(gps['Địa điểm tra GPS']) } : null,
      };
    });

    const stk = _extReadAll('HDMB', 'HD_STK').find((s) => String(s['Số hợp đồng']) === String(soHopDong));
    return {
      rungList,
      stk: stk ? {
        NguoiUyQuyen: _safeStr(stk['Họ và tên người được ủy quyền']),
        SoTKNhanTien: _safeStr(stk['Số TK nhận tiền']),
        NganHang: _safeStr(stk['Ngân hàng']),
        UyQuyenThanhToan: _safeStr(stk['Ủy quyền thanh toán']),
      } : null,
    };
  });
}

// ============ API: LÔ HÀNG MUA VÀO / XUẤT BÁN GẦN NHẤT ============
/** Chuyển Date object thành chuỗi text đơn giản trước khi gửi qua google.script.run —
 * tránh lỗi ngầm khi "đóng gói" payload lớn có nhiều cột kiểu Ngày/Giờ (lỗi Apps Script
 * khá phổ biến: client nhận về null dù server tính đúng dữ liệu). */
function _safeStr(v) {
  if (v === null || v === undefined) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    try { return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'); }
    catch (e) { return String(v); }
  }
  return v;
}
/** Chỉ lấy đúng các cột cần hiển thị (giảm kích thước payload đáng kể so với trả nguyên 29 cột).
 * Tải nhiều hơn (200 dòng) vì giờ mỗi dòng chỉ còn 9 trường gọn nhẹ — đủ cho phân trang 20 dòng/trang. */
function getLoHangGanDay(limit) {
  return _safe(() => {
    const rows = _extReadLastN('HOSOKEO', 'HoSoKeo_DN', limit || 200);
    return rows.map((r) => ({
      NgayNhap: _safeStr(r['Ngày nhập']),
      MaHopDong: _safeStr(r['Mã hợp đồng']),
      SoPhieuCan: _safeStr(r['SỐ PHIẾU CÂN']),
      SoBKLS: _safeStr(r['Số BKLS']),
      ChuRung: _safeStr(r['Họ và tên chủ rừng']),
      DiaChiRung: _safeStr(r['ĐỊA CHỈ RỪNG']),
      KhoiLuongTan: _safeStr(r['Khối lượng (Tấn)']),
      DonGia: _safeStr(r['Đơn giá']),
      ThanhToan: _safeStr(r['Thanh toán']),
    }));
  });
}
function getXuatHangGanDay(limit) {
  return _safe(() => {
    const rows = _extReadLastN('XUATHANG', 'NL_PC_XH', limit || 200);
    return rows.map((r) => ({
      SoPhieu: _safeStr(r['Số phiếu']),
      NgayCan1: _safeStr(r['Ngày giờ cân 1']),
      SoBKLS: _safeStr(r['Số BKLS']),
      KhoiLuongTan: _safeStr(r['Khối lượng (Tấn)']),
      KhoiLuongM3: _safeStr(r['Khối lượng (M3)']),
      DonViVanChuyen: _safeStr(r['Đơn vị vận chuyển']),
      SoTKHQ: _safeStr(r['SỐ TKHQ']),
      TauXuat: _safeStr(r['TÀU XUẤT']),
    }));
  });
}

// ============ API: ĐÁNH GIÁ HỢP ĐỒNG / RỦI RO RỪNG (dữ liệu MỚI) ============
function addDanhGiaHopDong(obj) {
  obj.NgayDanhGia = obj.NgayDanhGia || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return _ownAppend('DanhGiaHopDong', obj);
}
function addDanhGiaRuiRoRung(obj) {
  obj.NgayDanhGia = obj.NgayDanhGia || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return _ownAppend('DanhGiaRuiRoRung', obj);
}
function getDanhGiaRuiRoRungList() { return _safe(() => _ownReadAll('DanhGiaRuiRoRung')); }

// ============ API: GIÁM SÁT ĐỊNH KỲ / TIẾN ĐỘ / RỦI RO TRIỂN KHAI ============
function getGiamSatList() { return _safe(() => _ownReadAll('GiamSatDinhKy')); }
function addGiamSat(obj) {
  const rows = _ownReadAll('GiamSatDinhKy');
  obj.MaBaoCao = obj.MaBaoCao || 'GS-' + String(rows.length + 1).padStart(4, '0');
  return _ownAppend('GiamSatDinhKy', obj);
}
function updateTienDo(giaiDoan, ngayThucTe, danhGia, ghiChu) {
  const sh = _own('TienDoTrienKhai');
  const values = sh.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === giaiDoan) {
      sh.getRange(i + 1, 3).setValue(ngayThucTe);
      sh.getRange(i + 1, 5).setValue(danhGia);
      sh.getRange(i + 1, 6).setValue(ghiChu);
      return true;
    }
  }
  return false;
}
function getRuiRoTrienKhaiList() { return _safe(() => _ownReadAll('RuiRoTrienKhai')); }
function addRuiRoTrienKhai(obj) {
  const rows = _ownReadAll('RuiRoTrienKhai');
  obj.MaRuiRo = obj.MaRuiRo || 'RR-' + String(rows.length + 1).padStart(4, '0');
  return _safe(() => _ownAppend('RuiRoTrienKhai', obj));
}

// ============ API: KHẢO SÁT & THAM VẤN CÁC BÊN LIÊN QUAN (Bước 2, QT-FSC-01) ============
/** Đây là hồ sơ dạng Word (Báo cáo khảo sát, Biên bản giám sát khai thác/sau khai thác,
 * phỏng vấn UBND xã/Kiểm lâm/chủ rừng lân cận/người dân/công nhân/người vận chuyển) — KHÔNG
 * có sẵn dưới dạng bảng dữ liệu để đọc tự động như 5 sheet kia, nên chỉ theo dõi được bằng
 * cách người dùng tự đánh dấu Có/Chưa cho từng loại hồ sơ theo từng hợp đồng. */
function getKhaoSatList() { return _safe(() => _ownReadAll('KhaoSatThamVan')); }
function addKhaoSat(obj) {
  return _safe(() => _ownAppend('KhaoSatThamVan', obj));
}

/** Báo cáo khảo sát chi tiết — nhập TRỰC TIẾP trong webapp (không dùng Google Form nữa),
 * theo đúng 10 mục nội dung của mẫu "Báo cáo kết quả khảo sát" Word, có kèm ảnh khảo sát. */
function getBaoCaoKhaoSatList() { return _safe(() => _ownReadAll('BaoCaoKhaoSat')); }
function addBaoCaoKhaoSat(obj) {
  return _safe(() => _ownAppend('BaoCaoKhaoSat', obj));
}

/** Tự tạo (1 lần duy nhất) hoặc tái sử dụng thư mục Drive "Anh_Khao_Sat_FSC_HAK" để lưu ảnh khảo sát
 * tải lên từ webapp — dùng chung thư mục cha với mẫu Word (đã có quyền Drive từ trước). */
function _getOrCreateAnhKhaoSatFolder() {
  const props = PropertiesService.getScriptProperties();
  let folderId = props.getProperty('ANH_KHAO_SAT_FOLDER_ID');
  if (folderId) {
    try { return DriveApp.getFolderById(folderId); } catch (e) { /* thư mục cũ bị xóa, tạo lại */ }
  }
  const parent = DriveApp.getFolderById(MAU_WORD_FOLDER_ID);
  const folder = parent.createFolder('Anh_Khao_Sat_FSC_HAK');
  props.setProperty('ANH_KHAO_SAT_FOLDER_ID', folder.getId());
  return folder;
}
/** Nhận ảnh dạng base64 từ trình duyệt (đã resize nhỏ ở client trước khi gửi lên), lưu vào Drive,
 * trả về link xem trực tiếp. */
function uploadKhaoSatAnh(base64Data, fileName, mimeType) {
  return _safe(() => {
    const folder = _getOrCreateAnhKhaoSatFolder();
    const decoded = Utilities.base64Decode(base64Data);
    const blob = Utilities.newBlob(decoded, mimeType, fileName);
    const file = folder.createFile(blob);
    try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) { /* bỏ qua nếu không đủ quyền chia sẻ */ }
    return { url: file.getUrl(), name: fileName };
  });
}

/** Biên bản giám sát khai thác (trong/sau khai thác), theo đúng mẫu Word tương ứng */
function getBienBanGiamSatList() { return _safe(() => _ownReadAll('BienBanGiamSatKhaiThac')); }
function addBienBanGiamSat(obj) {
  return _safe(() => _ownAppend('BienBanGiamSatKhaiThac', obj));
}

/** Tham vấn từng bên liên quan (UBND xã, Kiểm lâm, chủ rừng lân cận, người dân, công nhân, người vận chuyển...) */
function getThamVanList() { return _safe(() => _ownReadAll('ThamVanBenLienQuan')); }
function addThamVan(obj) {
  return _safe(() => _ownAppend('ThamVanBenLienQuan', obj));
}

/** BM-FSC-01/02: Checklist tiêu chí thẩm định chủ rừng + kiểm tra trước ký hợp đồng */
function getTieuChiThamDinhList() { return _safe(() => _ownReadAll('TieuChiThamDinhHopDong')); }
function addTieuChiThamDinh(obj) {
  return _safe(() => _ownAppend('TieuChiThamDinhHopDong', obj));
}

/** BM-FSC-05: Lệnh điều động vận chuyển nội bộ */
function getLenhDieuDongList() { return _safe(() => _ownReadAll('LenhDieuDongVanChuyen')); }
function addLenhDieuDong(obj) {
  const rows = _ownReadAll('LenhDieuDongVanChuyen');
  obj.SoLenh = obj.SoLenh || 'LDD-' + String(rows.length + 1).padStart(4, '0');
  return _safe(() => _ownAppend('LenhDieuDongVanChuyen', obj));
}

/** BM-FSC-06: Phiếu theo dõi định mức tiêu hao chế biến */
function getTieuHaoCheBienList() { return _safe(() => _ownReadAll('TieuHaoCheBien')); }
function addTieuHaoCheBien(obj) {
  return _safe(() => _ownAppend('TieuHaoCheBien', obj));
}

/** BM-FSC-07 mở rộng: đọc thêm NL_DH_XB (đơn hàng xuất bán) — trước đây chưa dùng tới */
function getDonHangXuatBanList(limit) {
  return _safe(() => {
    const rows = _extReadLastN('XUATHANG', 'NL_DH_XB', limit || 200);
    return rows.map((r) => {
      const o = {};
      Object.keys(r).forEach((k) => { o[k] = _safeStr(r[k]); });
      return o;
    });
  });
}

/** BM-FSC-08: Sổ theo dõi nộp báo cáo định kỳ cơ quan Nhà nước (Mẫu 29, Mẫu 14) */
function getSoTheoDoiBaoCaoList() { return _safe(() => _ownReadAll('SoTheoDoiBaoCaoNhaNuoc')); }
function addSoTheoDoiBaoCao(obj) {
  return _safe(() => _ownAppend('SoTheoDoiBaoCaoNhaNuoc', obj));
}

/** BM-KH-02: Biên bản nghiệm thu hoàn thành giai đoạn triển khai */
function getNghiemThuGiaiDoanList() { return _safe(() => _ownReadAll('NghiemThuGiaiDoanTrienKhai')); }
function addNghiemThuGiaiDoan(obj) {
  return _safe(() => _ownAppend('NghiemThuGiaiDoanTrienKhai', obj));
}

// ============ ĐÁNH GIÁ NỘI BỘ HỆ THỐNG THẨM ĐỊNH DDS (theo FSC-STD-40-005 Ver.3.1) ============
/** Kế hoạch đánh giá nội bộ hằng năm — có công thức lấy mẫu y = 0.8 * sqrt(x) */
function getKeHoachDanhGiaList() { return _safe(() => _ownReadAll('KeHoachDanhGiaNoiBo')); }
function addKeHoachDanhGia(obj) {
  return _safe(() => _ownAppend('KeHoachDanhGiaNoiBo', obj));
}

/** BM.05: Phiếu yêu cầu hành động khắc phục (CAR) */
function getPhieuKhacPhucList() { return _safe(() => _ownReadAll('PhieuYeuCauKhacPhuc')); }
function addPhieuKhacPhuc(obj) {
  const rows = _ownReadAll('PhieuYeuCauKhacPhuc');
  obj.SoPhieu = obj.SoPhieu || 'BM05-' + String(rows.length + 1).padStart(4, '0');
  return _safe(() => _ownAppend('PhieuYeuCauKhacPhuc', obj));
}

/** BM.07: Biên bản xem xét hệ thống thẩm định DDS (Management Review) */
function getBienBanXemXetDDSList() { return _safe(() => _ownReadAll('BienBanXemXetHeThongDDS')); }
function addBienBanXemXetDDS(obj) {
  return _safe(() => _ownAppend('BienBanXemXetHeThongDDS', obj));
}

/** BM.08: Phê duyệt năng lực đánh giá viên nội bộ */
function getPheDuyetDanhGiaVienList() { return _safe(() => _ownReadAll('PheDuyetNangLucDanhGiaVien')); }
function addPheDuyetDanhGiaVien(obj) {
  return _safe(() => _ownAppend('PheDuyetNangLucDanhGiaVien', obj));
}

// ============ BÁO CÁO TỒN KHO & TÌNH TRẠNG KHẢO SÁT RỪNG ============
function _matchDateRange(dateVal, tuNgay, denNgay) {
  if (!tuNgay && !denNgay) return true;
  if (!dateVal) return false;
  const d = (Object.prototype.toString.call(dateVal) === '[object Date]') ? dateVal : new Date(dateVal);
  if (isNaN(d.getTime())) return false;
  if (tuNgay && d < new Date(tuNgay)) return false;
  if (denNgay && d > new Date(denNgay + 'T23:59:59')) return false;
  return true;
}
/** Tồn kho nguyên liệu = Tổng nhập (HoSoKeo_DN) - Tổng xuất (NL_PC_XH), lọc theo khoảng ngày (tùy chọn).
 * LƯU Ý: đây là tồn kho TỔNG, chưa lọc riêng theo tiêu chuẩn FSC/CW cụ thể vì cột "Nguồn gốc" trong
 * dữ liệu nguồn (PS/DT/QT...) chưa được xác nhận có phải mã phân loại chứng chỉ hay là mã khu vực địa lý. */
function getTonKhoBaoCao(tuNgay, denNgay) {
  return _safe(() => {
    const nhapSheet = _ext('HOSOKEO', 'HoSoKeo_DN');
    const xuatSheet = _ext('XUATHANG', 'NL_PC_XH');
    const nhapValues = nhapSheet.getDataRange().getValues();
    const xuatValues = xuatSheet.getDataRange().getValues();
    const nhapHeaders = nhapValues[0], xuatHeaders = xuatValues[0];
    const idxNgayNhap = nhapHeaders.indexOf('Ngày nhập');
    const idxKLNhap = nhapHeaders.indexOf('Khối lượng (Tấn)');
    const idxNguonGocNhap = nhapHeaders.indexOf('Nguồn gốc');
    const idxNgayXuat = xuatHeaders.indexOf('Ngày giờ cân 1');
    const idxKLXuat = xuatHeaders.indexOf('Khối lượng (Tấn)');

    let tongNhap = 0, soLoNhap = 0;
    const nguonGocSet = {};
    for (let i = 1; i < nhapValues.length; i++) {
      const row = nhapValues[i];
      if (!row[idxKLNhap] && row[idxKLNhap] !== 0) continue;
      if (!_matchDateRange(row[idxNgayNhap], tuNgay, denNgay)) continue;
      tongNhap += Number(row[idxKLNhap]) || 0;
      soLoNhap++;
      const ng = row[idxNguonGocNhap];
      if (ng) nguonGocSet[ng] = (nguonGocSet[ng] || 0) + (Number(row[idxKLNhap]) || 0);
    }
    let tongXuat = 0, soLoXuat = 0;
    for (let i = 1; i < xuatValues.length; i++) {
      const row = xuatValues[i];
      if (!row[idxKLXuat] && row[idxKLXuat] !== 0) continue;
      if (!_matchDateRange(row[idxNgayXuat], tuNgay, denNgay)) continue;
      tongXuat += Number(row[idxKLXuat]) || 0;
      soLoXuat++;
    }
    return {
      tongNhap: Math.round(tongNhap * 100) / 100,
      tongXuat: Math.round(tongXuat * 100) / 100,
      tonKho: Math.round((tongNhap - tongXuat) * 100) / 100,
      soLoNhap, soLoXuat,
      theoNguonGoc: Object.keys(nguonGocSet).map((k) => ({ nguonGoc: k, khoiLuong: Math.round(nguonGocSet[k] * 100) / 100 })),
    };
  });
}

/** Tình trạng khảo sát theo hợp đồng (đối chiếu HD_RUNG với sheet theo dõi KhaoSatThamVan) */
function getTinhTrangKhaoSatRung() {
  return _safe(() => {
    const rung = _extReadAll('HDMB', 'HD_RUNG');
    const khaoSat = _ownReadAll('KhaoSatThamVan');
    const ksMap = {};
    khaoSat.forEach((k) => { ksMap[String(k.SoHopDong)] = k; });
    const seen = {};
    const result = [];
    rung.forEach((r) => {
      const soHD = String(r.SoHopDong);
      if (seen[soHD]) return;
      seen[soHD] = true;
      const k = ksMap[soHD];
      result.push({
        SoHopDong: soHD,
        MaRung: _safeStr(r.MaRung),
        ChuRung: _safeStr(r.HoVaTenChuRung),
        DaKhaoSat: k && k.CoBaoCaoKhaoSat === 'Có' ? 'Có' : 'Chưa',
        NgayKhaoSat: k ? _safeStr(k.NgayKhaoSat) : '',
        NguoiKhaoSat: k ? _safeStr(k.NguoiKhaoSat) : '',
      });
    });
    return result;
  });
}

// ============ API: DASHBOARD TỔNG HỢP ============
function getDashboardData() {
  return _safe(() => _getDashboardDataInner());
}
function _getDashboardDataInner() {
  const rung = _extReadAll('HDMB', 'HD_RUNG');
  const kiemTra = _extReadAll('HDMB', 'BaoCao_KiemTra');
  const tongHop = _extReadAll('HDMB', 'TongHop_HopDong');
  const danhGia = _ownReadAll('DanhGiaHopDong');
  const ruiRoRung = _ownReadAll('DanhGiaRuiRoRung');
  const tienDo = _ownReadAll('TienDoTrienKhai');
  const giamSat = _ownReadAll('GiamSatDinhKy');
  const khaoSat = _ownReadAll('KhaoSatThamVan');
  const thamDinh = _ownReadAll('TieuChiThamDinhHopDong');

  // Tổng số hợp đồng: đếm theo SoHopDong duy nhất trong HD_RUNG (bảng gốc luôn có dữ liệu)
  const soHDSet = {};
  rung.forEach((r) => { soHDSet[String(r.SoHopDong)] = true; });
  const tongHD = Object.keys(soHDSet).length;

  const coDuLieuKiemTra = kiemTra.length > 0;
  const dayDu = kiemTra.filter((r) => String(r['Kết quả']).indexOf('Đầy đủ') > -1).length;
  const thieuHoSo = coDuLieuKiemTra ? (kiemTra.length - dayDu) : 0;

  const chenhLechCanhBao = tongHop.filter((r) => Math.abs(Number(r['Chênh lệch (%)']) || 0) > 10).length;

  // Một hợp đồng có thể được ghi nhận NHIỀU lượt đánh giá theo thời gian (tái đánh giá).
  // Chỉ lấy KẾT LUẬN GẦN NHẤT của mỗi hợp đồng (giống cách getHopDongList đang làm) để
  // không đếm 1 hợp đồng vào nhiều nhóm kết luận, và không "giấu" hợp đồng chưa từng đánh giá.
  const ketLuanMoiNhatMap = {};
  danhGia.forEach((r) => { ketLuanMoiNhatMap[String(r.MaHopDong)] = r.KetLuanGiamSat; });
  const ketLuanList = Object.values(ketLuanMoiNhatMap);
  const dat = ketLuanList.filter((kl) => kl === 'ĐẠT').length;
  const datCoDK = ketLuanList.filter((kl) => kl === 'ĐẠT CÓ ĐIỀU KIỆN').length;
  const khongDat = ketLuanList.filter((kl) => kl === 'KHÔNG ĐẠT').length;
  const chuaDanhGia = tongHD - Object.keys(ketLuanMoiNhatMap).length;

  const vungRuiRoDangKe = ruiRoRung.filter((r) => r.RuiRoCNRA === 'Đáng kể').length;
  const soHopDongCoKhaoSat = new Set(khaoSat.map((r) => String(r.SoHopDong))).size;
  const chuaKhaoSat = Math.max(tongHD - soHopDongCoKhaoSat, 0);
  const soHopDongDaThamDinh = new Set(thamDinh.map((r) => String(r.SoHopDong))).size;
  const chuaThamDinh = Math.max(tongHD - soHopDongDaThamDinh, 0);

  return {
    tongHD, dayDu, thieuHoSo,
    coDuLieuKiemTra,
    tyLeDayDu: coDuLieuKiemTra ? Math.round((dayDu / kiemTra.length) * 100) : null,
    chenhLechCanhBao,
    dat, datCoDK, khongDat, chuaDanhGia: Math.max(chuaDanhGia, 0),
    vungRuiRoDangKe,
    soBaoCaoGiamSat: giamSat.length,
    chuaKhaoSat,
    chuaThamDinh,
    tienDo,
  };
}
