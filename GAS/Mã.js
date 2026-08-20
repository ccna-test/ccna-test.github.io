const cauHinh = {
  data: {
    type: 'google sheet',
    fileId: '1nvGLMZwHsH0O213EWx52wUYVpmDgTk2zXgI1eec4-PI',
    sheetName: 'DATA_VI',
    cache: {
      startupLoad: 'all'
    }
  },

  cauTrucBaiKiemTra: {
    soCauHoi1Bai: 30,
    thoiGianLamBai: 30,
    xaoTronCauHoi: true
  },

  giaoDien: {
    brand: {
      colorPrimary: '#00a4de',
      colorSecondary: '#5dc8ed',
      iconUrl: 'https://www.citypng.com/public/uploads/preview/cisco-square-blue-logo-icon-png-735811696612218gzoiadfplh.png'
    }
  }
};

const TEN_CACHE = 'duLieuTracNghiemTiengViet';
const THOI_GIAN_CACHE = 21600;

function doGet() {
  return HtmlService
    .createHtmlOutputFromFile('index')
    .setTitle('Trac nghiem tieng Viet')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function layDuLieuCauHoi() {
  const cache = CacheService.getScriptCache();
  const duLieuCache = cache.get(TEN_CACHE);

  if (duLieuCache) {
    return JSON.parse(duLieuCache);
  }

  const duLieu = docDuLieuCauHoi();

  try {
    cache.put(TEN_CACHE, JSON.stringify(duLieu), THOI_GIAN_CACHE);
  } catch (loi) {
    // CacheService co gioi han kich thuoc; neu qua lon thi van tra du lieu truc tiep.
  }

  return duLieu;
}

function docDuLieuCauHoi() {
  const bangTinh = SpreadsheetApp.openById(cauHinh.data.fileId);
  const trangTinh = bangTinh.getSheetByName(cauHinh.data.sheetName);

  if (!trangTinh) {
    throw new Error(
      'Khong tim thay sheet "' + cauHinh.data.sheetName + '".'
    );
  }

  const soDong = trangTinh.getLastRow();

  if (soDong < 2) {
    return [];
  }

  const duLieu = trangTinh
    .getRange(1, 1, soDong, 7)
    .getDisplayValues();

  const danhSachCauHoi = [];

  for (let chiSo = 1; chiSo < duLieu.length; chiSo++) {
    const dong = duLieu[chiSo];

    const noiDung = String(dong[0] || '').trim();
    const luaChon = [
      String(dong[1] || '').trim(),
      String(dong[2] || '').trim(),
      String(dong[3] || '').trim(),
      String(dong[4] || '').trim()
    ];
    const dapAn = Number(String(dong[5] || '').trim());
    const giaiThich = String(dong[6] || '').trim();

    if (!noiDung) {
      continue;
    }

    if (
      !Number.isInteger(dapAn) ||
      dapAn < 0 ||
      dapAn > 3
    ) {
      continue;
    }

    danhSachCauHoi.push({
      noiDung: noiDung,
      luaChon: luaChon,
      dapAn: dapAn,
      giaiThich: giaiThich
    });
  }

  return danhSachCauHoi;
}

function xoaCacheCauHoi() {
  CacheService.getScriptCache().remove(TEN_CACHE);
  return 'Da xoa cache';
}

function kiemTraDuLieu() {
  const duLieu = docDuLieuCauHoi();

  return {
    soCauHoi: duLieu.length,
    soCauCan: cauHinh.cauTrucBaiKiemTra.soCauHoi1Bai,
    duDieuKien:
      duLieu.length >= cauHinh.cauTrucBaiKiemTra.soCauHoi1Bai
  };
}