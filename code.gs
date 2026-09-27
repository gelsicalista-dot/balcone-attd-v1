/**
 * THE BALCONE SUITES & RESORT - Backend API Engine
 * File: code.gs
 * Author: ZettBOT Assistant by Zettbos
 */

const CONFIG = {
  SHEET_KARYAWAN: 'Tabel_Karyawan',
  SHEET_SHIFT: 'Tabel_Shift',
  SHEET_DEPARTEMEN: 'Tabel_Departemen',
  SHEET_ABSENSI: 'Tabel_Absensi',
  SHEET_IZIN: 'Tabel_Izin_Cuti',
  SHEET_KPI: 'Tabel_KPI',
  SHEET_PENGUMUMAN: 'Tabel_Pengumuman',
  SHEET_PENGATURAN: 'Tabel_Pengaturan',
  SHEET_ROSTER: 'Tabel_Roster',
  SHEET_EDIT_PROFIL: 'Tabel_Edit_Profil',
  SHEET_GAJI_MASTER: 'Tabel_Gaji_Master',
  SHEET_PAYROLL_BULANAN: 'Tabel_Payroll_Bulanan',
  SHEET_PINJAMAN: 'Tabel_Pinjaman_Karyawan'
};

function doGet(e) {
  const template = HtmlService.createTemplateFromFile('index');
  return template.evaluate()
    .setTitle('THE BALCONE SUITES & RESORT - HR & Payroll System')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Helper Function untuk merender file modul komponen terpisah (HTML/CSS/JS)
 * Menggunakan sintaks scriptlet: <?!= include('namafile'); ?>
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return ContentService.createTextOutput(JSON.stringify({ success: false, message: 'Request payload kosong.' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const requestData = JSON.parse(e.postData.contents);
    const action = requestData.action;
    const payload = requestData.payload || [];
    
    let result = null;
    
    if (action === 'loginUser') {
      result = loginUser(payload[0], payload[1]);
    } else if (action === 'registerKaryawan') {
      result = registerKaryawan(payload[0]);
    } else if (action === 'approveKaryawan') {
      result = approveKaryawan(payload[0], payload[1]);
    } else if (action === 'setEmployeeResign') {
      result = setEmployeeResign(payload[0], payload[1], payload[2]);
    } else if (action === 'processAbsensi') {
      result = processAbsensi(payload[0], payload[1], payload[2], payload[3], payload[4]);
    } else if (action === 'getKaryawanDashboard') {
      result = getKaryawanDashboard(payload[0]);
    } else if (action === 'getHODDashboard') {
      result = getHODDashboard(payload[0], payload[1]);
    } else if (action === 'getAdminDashboard') {
      result = getAdminDashboard();
    } else if (action === 'submitIzinCuti') {
      result = submitIzinCuti(payload[0], payload[1], payload[2], payload[3], payload[4], payload[5]);
    } else if (action === 'approveIzinCuti') {
      result = approveIzinCuti(payload[0], payload[1], payload[2]);
    } else if (action === 'calculateMonthlyKPI') {
      result = calculateMonthlyKPI(payload[0]);
    } else if (action === 'getGlobalAttendanceList') {
      result = getGlobalAttendanceList(payload[0], payload[1], payload[2], payload[3]);
    } else if (action === 'savePengumuman') {
      result = savePengumuman(payload[0], payload[1], payload[2], payload[3], payload[4]);
    } else if (action === 'deletePengumuman') {
      result = deletePengumuman(payload[0]);
    } else if (action === 'getShiftList') {
      result = getShiftList();
    } else if (action === 'saveShift') {
      result = saveShift(payload[0], payload[1], payload[2], payload[3], payload[4], payload[5]);
    } else if (action === 'deleteShift') {
      result = deleteShift(payload[0]);
    } else if (action === 'getDepartemenList') {
      result = getDepartemenList();
    } else if (action === 'saveDepartemen') {
      result = saveDepartemen(payload[0], payload[1]);
    } else if (action === 'deleteDepartemen') {
      result = deleteDepartemen(payload[0]);
    } else if (action === 'saveOfficeSettings') {
      result = saveOfficeSettings(payload[0], payload[1], payload[2], payload[3], payload[4], payload[5]);
    } else if (action === 'getOfficeSettings') {
      result = getOfficeSettings();
    } else if (action === 'generatePDFReport') {
      result = generatePDFReport(payload[0]);
    } else if (action === 'getRosterData') {
      result = getRosterData(payload[0], payload[1]);
    } else if (action === 'saveBulkRoster') {
      result = saveBulkRoster(payload[0]);
    } else if (action === 'submitEditProfil') {
      result = submitEditProfil(payload[0], payload[1]);
    } else if (action === 'approveEditProfil') {
      result = approveEditProfil(payload[0], payload[1]);
    } else if (action === 'verifyPayrollPIN') {
      result = verifyPayrollPIN(payload[0]);
    } else if (action === 'getSalaryMasterList') {
      result = getSalaryMasterList();
    } else if (action === 'saveSalaryMaster') {
      result = saveSalaryMaster(payload[0], payload[1]);
    } else if (action === 'calculateMonthlyPayroll') {
      result = calculateMonthlyPayroll(payload[0]);
    } else if (action === 'savePayrollRun') {
      result = savePayrollRun(payload[0], payload[1]);
    } else if (action === 'getPersonalPayslip') {
      result = getPersonalPayslip(payload[0], payload[1]);
    } else if (action === 'getLoansList') {
      result = getLoansList();
    } else if (action === 'saveLoan') {
      result = saveLoan(payload[0], payload[1], payload[2], payload[3]);
    } else {
      result = { success: false, message: 'Action API tidak dikenal.' };
    }

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    const errorRes = { success: false, message: 'Server API Error: ' + err.toString() };
    return ContentService.createTextOutput(JSON.stringify(errorRes))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getSheetData(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  return sheet.getDataRange().getDisplayValues();
}

function getSheetDataAsObjects(sheetName) {
  const data = getSheetData(sheetName);
  if (data.length <= 1) return [];
  const headers = data[0];
  const rows = data.slice(1);
  return rows.map(row => {
    let obj = {};
    headers.forEach((header, index) => {
      obj[header] = row[index] || '';
    });
    return obj;
  });
}

function parseDateStrToDate(dateStr) {
  if (!dateStr) return null;
  dateStr = dateStr.toString().trim();
  let year, month, day;
  if (dateStr.includes('-')) {
    const parts = dateStr.split('-');
    if (parts[0].length === 4) {
      year = parseInt(parts[0], 10);
      month = parseInt(parts[1], 10) - 1;
      day = parseInt(parts[2], 10);
    } else {
      day = parseInt(parts[0], 10);
      month = parseInt(parts[1], 10) - 1;
      year = parseInt(parts[2], 10);
    }
  } else if (dateStr.includes('/')) {
    const parts = dateStr.split('/');
    if (parts[0].length === 4) {
      year = parseInt(parts[0], 10);
      month = parseInt(parts[1], 10) - 1;
      day = parseInt(parts[2], 10);
    } else {
      day = parseInt(parts[0], 10);
      month = parseInt(parts[1], 10) - 1;
      year = parseInt(parts[2], 10);
    }
  } else {
    return new Date(dateStr);
  }
  return new Date(year, month, day);
}

function getCutoffRange(monthYearStr) {
  let year, month;
  if (monthYearStr && monthYearStr.includes('-')) {
    const parts = monthYearStr.split('-');
    year = parseInt(parts[0], 10);
    month = parseInt(parts[1], 10);
  } else if (monthYearStr && monthYearStr.includes('/')) {
    const parts = monthYearStr.split('/');
    month = parseInt(parts[0], 10);
    year = parseInt(parts[1], 10);
  } else {
    const now = new Date();
    year = now.getFullYear();
    month = now.getMonth() + 1;
  }

  let prevYear = year;
  let prevMonth = month - 1;
  if (prevMonth === 0) {
    prevMonth = 12;
    prevYear = year - 1;
  }

  const startDate = new Date(prevYear, prevMonth - 1, 26, 0, 0, 0);
  const endDate = new Date(year, month - 1, 25, 23, 59, 59);

  return {
    startDate: startDate,
    endDate: endDate,
    startIsoStr: Utilities.formatDate(startDate, 'Asia/Jakarta', 'yyyy-MM-dd'),
    endIsoStr: Utilities.formatDate(endDate, 'Asia/Jakarta', 'yyyy-MM-dd')
  };
}

function haversineDistance(lat1, lon1, lat2, lon2) {
  const R = 6371000;
  const dLat = (lat2 - lat1) * Math.PI / 180;
  const dLon = (lon2 - lon1) * Math.PI / 180;
  const a = 
    Math.sin(dLat / 2) * Math.sin(dLat / 2) +
    Math.cos(lat1 * Math.PI / 180) * Math.cos(lat2 * Math.PI / 180) * 
    Math.sin(dLon / 2) * Math.sin(dLon / 2);
  const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
  return Math.round(R * c);
}

function generateNIK() {
  const employees = getSheetDataAsObjects(CONFIG.SHEET_KARYAWAN);
  let maxSeq = 0;
  employees.forEach(e => {
    if (e.nik && e.nik.indexOf('NIK-') === 0) {
      const seqStr = e.nik.replace('NIK-', '');
      const seqNum = parseInt(seqStr, 10);
      if (!isNaN(seqNum) && seqNum > maxSeq) {
        maxSeq = seqNum;
      }
    }
  });
  return 'NIK-' + (maxSeq + 1).toString().padStart(4, '0');
}

function generateDailyId(prefix, sheetName) {
  const today = new Date();
  const dateStr = Utilities.formatDate(today, 'Asia/Jakarta', 'yyyyMMdd');
  const fullPrefix = prefix + '-' + dateStr + '-';
  const data = getSheetData(sheetName);
  let maxSeq = 0;
  
  for (let i = 1; i < data.length; i++) {
    const id = data[i][0];
    if (id && id.indexOf(fullPrefix) === 0) {
      const seqNum = parseInt(id.replace(fullPrefix, ''), 10);
      if (!isNaN(seqNum) && seqNum > maxSeq) maxSeq = seqNum;
    }
  }
  return fullPrefix + (maxSeq + 1).toString().padStart(4, '0');
}

function sanitizePhone(phone) {
  if (!phone) return '';
  let str = phone.toString().trim();
  if (/[a-zA-Z]/.test(str)) {
    return str;
  }
  let cleaned = str.replace(/[^0-9]/g, '');
  if (cleaned.startsWith('62')) {
    cleaned = '0' + cleaned.substring(2);
  } else if (cleaned.length > 0 && !cleaned.startsWith('0')) {
    cleaned = '0' + cleaned;
  }
  return cleaned;
}

function uploadFotoToDrive(base64Data, filename) {
  try {
    if (!base64Data || typeof base64Data !== 'string' || !base64Data.includes('base64,')) {
      return 'https://images.unsplash.com/photo-1535713875002-d1d0cf377fde?w=150';
    }

    const parts = base64Data.split('base64,');
    const header = parts[0];
    const rawBase64 = parts[1];
    
    let contentType = 'image/jpeg';
    if (header.includes('image/png')) contentType = 'image/png';
    else if (header.includes('image/webp')) contentType = 'image/webp';

    const decoded = Utilities.base64Decode(rawBase64);
    const blob = Utilities.newBlob(decoded, contentType, filename);

    const folderName = 'Balcone_Foto_Karyawan';
    const folders = DriveApp.getFoldersByName(folderName);
    let folder;
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      folder = DriveApp.createFolder(folderName);
      folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    }

    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    const fileId = file.getId();
    return 'https://lh3.googleusercontent.com/d/' + fileId;
  } catch (err) {
    Logger.log('Gagal upload foto ke Drive: ' + err.toString());
    return 'https://images.unsplash.com/photo-1535713875002-d1d0cf377fde?w=150';
  }
}

function loginUser(noTelp, password) {
  try {
    const cleanedInputPhone = sanitizePhone(noTelp);
    const employees = getSheetDataAsObjects(CONFIG.SHEET_KARYAWAN);
    
    const user = employees.find(e => {
      const dbPhone = sanitizePhone(e.no_telp);
      return dbPhone.toLowerCase() === cleanedInputPhone.toLowerCase() && e.password === password;
    });

    if (!user) {
      return { success: false, message: 'No. Telepon / Username atau Password salah.' };
    }

    if (user.status_akun === 'Pending') {
      return { 
        success: false, 
        message: 'Akun Anda sedang menunggu persetujuan dari Admin HR. Silakan hubungi Tim HRD.' 
      };
    } else if (user.status_akun === 'Rejected') {
      return { 
        success: false, 
        message: 'Pengajuan akun Anda ditolak oleh Admin HR.' 
      };
    }

    return {
      success: true,
      user: {
        nik: user.nik,
        nama: user.nama,
        no_telp: user.no_telp,
        no_identitas: user.no_identitas || '',
        alamat: user.alamat || '',
        status_kawin: user.status_kawin || 'TK/0',
        bank: user.bank || 'Mandiri',
        no_rekening: user.no_rekening || '',
        role: user.role,
        departemen: user.departemen,
        jabatan: user.jabatan,
        status_karyawan: user.status_karyawan,
        status_kerja: user.status_kerja || 'Aktif',
        tgl_resign: user.tgl_resign || '',
        tgl_lahir: user.tgl_lahir || '-',
        foto: user.foto || 'https://images.unsplash.com/photo-1535713875002-d1d0cf377fde?w=150',
        kuota_cuti_tahunan: user.kuota_cuti_tahunan,
        sisa_cuti_tahunan: user.sisa_cuti_tahunan
      }
    };
  } catch (err) {
    return { success: false, message: 'Terjadi kesalahan sistem: ' + err.toString() };
  }
}

function registerKaryawan(formData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetKaryawan = ss.getSheetByName(CONFIG.SHEET_KARYAWAN);
    const employees = getSheetDataAsObjects(CONFIG.SHEET_KARYAWAN);
    
    const formattedPhone = sanitizePhone(formData.no_telp);

    const isExist = employees.some(e => sanitizePhone(e.no_telp) === formattedPhone);
    if (isExist) {
      return { success: false, message: 'Nomor Telepon ' + formattedPhone + ' sudah terdaftar!' };
    }

    const newNIK = generateNIK();
    const phoneToStore = "'" + formattedPhone;

    let fotoProfilUrl = 'https://images.unsplash.com/photo-1535713875002-d1d0cf377fde?w=150';
    if (formData.foto && formData.foto.includes('base64,')) {
      fotoProfilUrl = uploadFotoToDrive(formData.foto, newNIK + '_Foto');
    }

    sheetKaryawan.appendRow([
      newNIK,
      formData.nama,
      phoneToStore,
      formData.password,
      formData.role || 'Karyawan',
      formData.status_karyawan,
      formData.no_identitas,
      formData.alamat,
      formData.status_kawin,
      formData.departemen,
      formData.jabatan,
      formData.tgl_masuk,
      formData.tgl_lahir,
      formData.bank,
      formData.no_rekening,
      fotoProfilUrl,
      '12',
      '12',
      'Pending',
      'Aktif',
      ''
    ]);

    const sheetGajiMaster = ss.getSheetByName(CONFIG.SHEET_GAJI_MASTER);
    sheetGajiMaster.appendRow([newNIK, '3000000', '0', '0', '0', '1000', 'Ya', 'Ya', 'Otomatis']);

    SpreadsheetApp.flush();
    return {
      success: true,
      nik: newNIK,
      message: 'Pendaftaran Berhasil! NIK Anda: ' + newNIK + '. Akun Anda menunggu persetujuan Admin HR.'
    };
  } catch (err) {
    return { success: false, message: 'Gagal melakukan pendaftaran: ' + err.toString() };
  }
}

function setEmployeeResign(nik, tglResign, statusKerja) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_KARYAWAN);
    const employees = getSheetDataAsObjects(CONFIG.SHEET_KARYAWAN);
    const index = employees.findIndex(e => e.nik === nik);

    if (index === -1) return { success: false, message: 'Data karyawan tidak ditemukan.' };

    const rowNum = index + 2;
    sheet.getRange(rowNum, 20).setValue(statusKerja || 'Resign');
    sheet.getRange(rowNum, 21).setValue(tglResign || '');

    SpreadsheetApp.flush();
    return { 
      success: true, 
      message: 'Status kerja karyawan ' + nik + ' diperbarui menjadi: ' + (statusKerja || 'Resign') + ' (Tgl: ' + (tglResign || '-') + ')' 
    };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function verifyPayrollPIN(pinInput) {
  try {
    const settings = getSheetDataAsObjects(CONFIG.SHEET_PENGATURAN)[0] || {};
    const validPin = settings.pin_payroll || '123456';

    if (pinInput === validPin) {
      return { success: true, message: 'PIN Verifikasi Payroll Valid!' };
    } else {
      return { success: false, message: 'PIN Otorisasi Payroll Salah!' };
    }
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function getSalaryMasterList() {
  try {
    const employees = getSheetDataAsObjects(CONFIG.SHEET_KARYAWAN).filter(k => k.status_akun === 'Approved' && k.role !== 'Admin HR');
    const salaryMasters = getSheetDataAsObjects(CONFIG.SHEET_GAJI_MASTER);

    const result = employees.map(emp => {
      const sal = salaryMasters.find(s => s.nik === emp.nik) || {};
      return {
        nik: emp.nik,
        nama: emp.nama,
        departemen: emp.departemen,
        jabatan: emp.jabatan,
        status_karyawan: emp.status_karyawan,
        status_kerja: emp.status_kerja || 'Aktif',
        tgl_resign: emp.tgl_resign || '',
        gaji_pokok: sal.gaji_pokok || '3000000',
        tunjangan_jabatan: sal.tunjangan_jabatan || '0',
        tunjangan_makan: sal.tunjangan_makan || '0',
        tunjangan_transport: sal.tunjangan_transport || '0',
        rate_denda_per_menit: sal.rate_denda_per_menit || '1000',
        bpjs_tk_aktif: sal.bpjs_tk_aktif || 'Ya',
        bpjs_kes_aktif: sal.bpjs_kes_aktif || 'Ya',
        mode_prorate: sal.mode_prorate || 'Otomatis'
      };
    });

    return { success: true, data: result };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function saveSalaryMaster(nik, salaryData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_GAJI_MASTER);
    const masterList = getSheetDataAsObjects(CONFIG.SHEET_GAJI_MASTER);

    const idx = masterList.findIndex(m => m.nik === nik);
    if (idx !== -1) {
      const row = idx + 2;
      sheet.getRange(row, 2).setValue(salaryData.gaji_pokok || '0');
      sheet.getRange(row, 3).setValue(salaryData.tunjangan_jabatan || '0');
      sheet.getRange(row, 4).setValue(salaryData.tunjangan_makan || '0');
      sheet.getRange(row, 5).setValue(salaryData.tunjangan_transport || '0');
      sheet.getRange(row, 6).setValue(salaryData.rate_denda_per_menit || '1000');
      sheet.getRange(row, 7).setValue(salaryData.bpjs_tk_aktif || 'Ya');
      sheet.getRange(row, 8).setValue(salaryData.bpjs_kes_aktif || 'Ya');
      sheet.getRange(row, 9).setValue(salaryData.mode_prorate || 'Otomatis');
    } else {
      sheet.appendRow([
        nik,
        salaryData.gaji_pokok || '0',
        salaryData.tunjangan_jabatan || '0',
        salaryData.tunjangan_makan || '0',
        salaryData.tunjangan_transport || '0',
        salaryData.rate_denda_per_menit || '1000',
        salaryData.bpjs_tk_aktif || 'Ya',
        salaryData.bpjs_kes_aktif || 'Ya',
        salaryData.mode_prorate || 'Otomatis'
      ]);
    }

    SpreadsheetApp.flush();
    return { success: true, message: 'Struktur gaji master karyawan ' + nik + ' berhasil diperbarui!' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function getLoansList() {
  try {
    const loans = getSheetDataAsObjects(CONFIG.SHEET_PINJAMAN);
    const employees = getSheetDataAsObjects(CONFIG.SHEET_KARYAWAN);

    const result = loans.map(l => {
      const emp = employees.find(e => e.nik === l.nik) || {};
      return {
        ...l,
        nama: emp.nama || 'N/A',
        departemen: emp.departemen || '-'
      };
    }).reverse();

    return { success: true, data: result };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function saveLoan(nik, totalPinjaman, cicilanPerBulan, tanggalPinjam) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_PINJAMAN);
    const newId = 'LNM-' + Date.now().toString().slice(-6);

    sheet.appendRow([
      newId,
      nik,
      tanggalPinjam || Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd'),
      totalPinjaman,
      cicilanPerBulan,
      totalPinjaman,
      'Belum Lunas'
    ]);

    SpreadsheetApp.flush();
    return { success: true, message: 'Pinjaman karyawan baru berhasil dicatat!' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function calculateMonthlyPayroll(bulanTahun) {
  try {
    const cutOff = getCutoffRange(bulanTahun);
    let karyawan = getSheetDataAsObjects(CONFIG.SHEET_KARYAWAN).filter(k => k.status_akun === 'Approved' && k.role !== 'Admin HR');
    
    karyawan = karyawan.filter(emp => {
      if (!emp.tgl_resign || emp.tgl_resign === '') return true;
      return emp.tgl_resign >= cutOff.startIsoStr;
    });

    const salaryMasters = getSheetDataAsObjects(CONFIG.SHEET_GAJI_MASTER);
    const absensi = getSheetDataAsObjects(CONFIG.SHEET_ABSENSI);
    const loans = getSheetDataAsObjects(CONFIG.SHEET_PINJAMAN);
    const existingPayroll = getSheetDataAsObjects(CONFIG.SHEET_PAYROLL_BULANAN);

    const payrollList = karyawan.map(emp => {
      const existingRecord = existingPayroll.find(p => p.nik === emp.nik && p.bulan_tahun === bulanTahun);

      if (existingRecord && existingRecord.status_bayar === 'Paid / Final') {
        return {
          nik: emp.nik,
          nama: emp.nama,
          no_identitas: emp.no_identitas || '-',
          departemen: emp.departemen,
          jabatan: emp.jabatan,
          status_karyawan: emp.status_karyawan,
          bank: emp.bank || 'Mandiri',
          no_rekening: emp.no_rekening || '-',
          bulan_tahun: bulanTahun,
          total_hadir: parseInt(existingRecord.total_hadir || '0', 10),
          total_menit_terlambat: parseInt(existingRecord.total_menit_terlambat || '0', 10),
          gaji_pokok: parseFloat(existingRecord.gaji_pokok || '0'),
          tunjangan: parseFloat(existingRecord.tunjangan || '0'),
          upah_lembur: parseFloat(existingRecord.upah_lembur || '0'),
          service_charge: parseFloat(existingRecord.service_charge || '0'),
          bpjs_tk: parseFloat(existingRecord.bpjs_tk || '0'),
          bpjs_kes: parseFloat(existingRecord.bpjs_kes || '0'),
          potongan_pinjaman: parseFloat(existingRecord.potongan_pinjaman || '0'),
          denda_terlambat: parseFloat(existingRecord.denda_terlambat || '0'),
          thp_bersih: parseFloat(existingRecord.thp_bersih || '0'),
          status_bayar: existingRecord.status_bayar,
          tanggal_transfer: existingRecord.tanggal_transfer || '-'
        };
      }

      const salMaster = salaryMasters.find(s => s.nik === emp.nik) || {
        gaji_pokok: '3000000', tunjangan_jabatan: '0', tunjangan_makan: '0', tunjangan_transport: '0',
        rate_denda_per_menit: '1000', bpjs_tk_aktif: 'Ya', bpjs_kes_aktif: 'Ya', mode_prorate: 'Otomatis'
      };

      const empAbsensi = absensi.filter(a => {
        if (a.nik !== emp.nik || !a.tanggal) return false;
        const parts = a.tanggal.split('/');
        if (parts.length === 3) {
          const recDate = new Date(parseInt(parts[2], 10), parseInt(parts[1], 10) - 1, parseInt(parts[0], 10));
          return recDate >= cutOff.startDate && recDate <= cutOff.endDate;
        }
        return false;
      });

      const totalHadir = empAbsensi.filter(a => a.status === 'Tepat Waktu' || a.status === 'Terlambat').length;
      let totalMenitTerlambat = 0;
      empAbsensi.forEach(a => {
        totalMenitTerlambat += parseInt(a.keterlambatan_menit || '0', 10);
      });

      const baseGajiPokok = parseFloat(salMaster.gaji_pokok || '0');
      const tunjJab = parseFloat(salMaster.tunjangan_jabatan || '0');
      const tunjMakan = parseFloat(salMaster.tunjangan_makan || '0');
      const tunjTransport = parseFloat(salMaster.tunjangan_transport || '0');
      const baseTotalTunjangan = tunjJab + tunjMakan + tunjTransport;

      const modeProrate = salMaster.mode_prorate || 'Otomatis';
      let isProrateApplied = false;

      if (modeProrate === 'Paksa Prorate') {
        isProrateApplied = true;
      } else if (modeProrate === 'Paksa Full (100%)') {
        isProrateApplied = false;
      } else {
        const isNewEmployeeInCutoff = emp.tgl_masuk && emp.tgl_masuk > cutOff.startIsoStr && emp.tgl_masuk <= cutOff.endIsoStr;
        const isResignInCutoff = emp.tgl_resign && emp.tgl_resign >= cutOff.startIsoStr && emp.tgl_resign <= cutOff.endIsoStr;
        if (isNewEmployeeInCutoff || isResignInCutoff) {
          isProrateApplied = true;
        }
      }

      const targetHariKerja = 26;
      const prorateRatio = isProrateApplied ? Math.min(1, totalHadir / targetHariKerja) : 1;

      let gantiGajiPokok = baseGajiPokok;
      let totalTunjangan = baseTotalTunjangan;

      if (emp.status_karyawan === 'DW' || emp.status_karyawan === 'Casual') {
        gantiGajiPokok = baseGajiPokok * totalHadir;
        totalTunjangan = baseTotalTunjangan;
      } else if (isProrateApplied) {
        gantiGajiPokok = Math.round(baseGajiPokok * prorateRatio);
        totalTunjangan = Math.round(baseTotalTunjangan * prorateRatio);
      }

      const upahLembur = 0;
      const serviceCharge = 0;
      const grossEarnings = gantiGajiPokok + totalTunjangan + upahLembur + serviceCharge;

      let potBPJSTK = 0;
      let potBPJSKes = 0;

      if (salMaster.bpjs_tk_aktif === 'Ya') {
        if (emp.status_karyawan === 'DW' || emp.status_karyawan === 'Casual') {
          potBPJSTK = Math.round(grossEarnings * 0.0084);
        } else {
          potBPJSTK = Math.round(gantiGajiPokok * 0.03);
        }
      }

      if (salMaster.bpjs_kes_aktif === 'Ya') {
        if (emp.status_karyawan === 'DW' || emp.status_karyawan === 'Casual') {
          potBPJSKes = 0;
        } else {
          potBPJSKes = Math.round(gantiGajiPokok * 0.01);
        }
      }

      const rateDenda = parseFloat(salMaster.rate_denda_per_menit || '1000');
      const dendaTerlambat = Math.round(totalMenitTerlambat * rateDenda);

      let potPinjaman = 0;
      if (existingRecord && existingRecord.potongan_pinjaman !== undefined && existingRecord.potongan_pinjaman !== '') {
        potPinjaman = parseFloat(existingRecord.potongan_pinjaman || '0');
      } else {
        const activeLoan = loans.find(l => l.nik === emp.nik && l.status_lunas === 'Belum Lunas');
        if (activeLoan) {
          potPinjaman = Math.min(parseFloat(activeLoan.sisa_pinjaman || '0'), parseFloat(activeLoan.cicilan_per_bulan || '0'));
        }
      }

      const totalPotongan = potBPJSTK + potBPJSKes + dendaTerlambat + potPinjaman;
      const thpBersih = Math.max(0, Math.round(grossEarnings - totalPotongan));

      return {
        nik: emp.nik,
        nama: emp.nama,
        no_identitas: emp.no_identitas || '-',
        departemen: emp.departemen,
        jabatan: emp.jabatan,
        status_karyawan: emp.status_karyawan,
        status_kerja: emp.status_kerja || 'Aktif',
        bank: emp.bank || 'Mandiri',
        no_rekening: emp.no_rekening || '-',
        bulan_tahun: bulanTahun,
        total_hadir: totalHadir,
        total_menit_terlambat: totalMenitTerlambat,
        gaji_pokok: gantiGajiPokok,
        tunjangan: totalTunjangan,
        upah_lembur: upahLembur,
        service_charge: serviceCharge,
        bpjs_tk: potBPJSTK,
        bpjs_kes: potBPJSKes,
        potongan_pinjaman: potPinjaman,
        denda_terlambat: dendaTerlambat,
        thp_bersih: thpBersih,
        status_bayar: existingRecord ? existingRecord.status_bayar : 'Draft',
        tanggal_transfer: existingRecord ? existingRecord.tanggal_transfer : '-'
      };
    });

    return { success: true, data: payrollList };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function savePayrollRun(bulanTahun, payrollItems) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetPayroll = ss.getSheetByName(CONFIG.SHEET_PAYROLL_BULANAN);
    const existingPayroll = getSheetDataAsObjects(CONFIG.SHEET_PAYROLL_BULANAN);
    const sheetLoans = ss.getSheetByName(CONFIG.SHEET_PINJAMAN);
    const loans = getSheetDataAsObjects(CONFIG.SHEET_PINJAMAN);

    const todayStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm');

    payrollItems.forEach(item => {
      const idx = existingPayroll.findIndex(p => p.nik === item.nik && p.bulan_tahun === bulanTahun);
      
      if (idx !== -1) {
        const row = idx + 2;
        sheetPayroll.getRange(row, 4).setValue(item.total_hadir.toString());
        sheetPayroll.getRange(row, 5).setValue(item.total_menit_terlambat.toString());
        sheetPayroll.getRange(row, 6).setValue(item.gaji_pokok.toString());
        sheetPayroll.getRange(row, 7).setValue(item.tunjangan.toString());
        sheetPayroll.getRange(row, 8).setValue((item.upah_lembur || 0).toString());
        sheetPayroll.getRange(row, 9).setValue((item.service_charge || 0).toString());
        sheetPayroll.getRange(row, 10).setValue(item.bpjs_tk.toString());
        sheetPayroll.getRange(row, 11).setValue(item.bpjs_kes.toString());
        sheetPayroll.getRange(row, 12).setValue(item.potongan_pinjaman.toString());
        sheetPayroll.getRange(row, 13).setValue(item.denda_terlambat.toString());
        sheetPayroll.getRange(row, 14).setValue(item.thp_bersih.toString());
        sheetPayroll.getRange(row, 15).setValue(todayStr);
        sheetPayroll.getRange(row, 16).setValue('Paid / Final');
      } else {
        const newId = 'PYR-' + Date.now().toString().slice(-6) + '-' + Math.floor(Math.random() * 100);
        sheetPayroll.appendRow([
          newId,
          item.nik,
          bulanTahun,
          item.total_hadir.toString(),
          item.total_menit_terlambat.toString(),
          item.gaji_pokok.toString(),
          item.tunjangan.toString(),
          (item.upah_lembur || 0).toString(),
          (item.service_charge || 0).toString(),
          item.bpjs_tk.toString(),
          item.bpjs_kes.toString(),
          item.potongan_pinjaman.toString(),
          item.denda_terlambat.toString(),
          item.thp_bersih.toString(),
          todayStr,
          'Paid / Final'
        ]);
      }

      if (item.potongan_pinjaman > 0) {
        const loanIdx = loans.findIndex(l => l.nik === item.nik && l.status_lunas === 'Belum Lunas');
        if (loanIdx !== -1) {
          const currentSisa = parseFloat(loans[loanIdx].sisa_pinjaman || '0');
          const newSisa = Math.max(0, currentSisa - item.potongan_pinjaman);
          sheetLoans.getRange(loanIdx + 2, 6).setValue(newSisa.toString());
          if (newSisa === 0) {
            sheetLoans.getRange(loanIdx + 2, 7).setValue('Lunas');
          }
        }
      }
    });

    SpreadsheetApp.flush();
    return { success: true, message: 'Proses Payroll Siklus Cut-Off ' + bulanTahun + ' Berhasil Difinalisasi & Disimpan!' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function getPersonalPayslip(nik, bulanTahun) {
  try {
    const payrollList = getSheetDataAsObjects(CONFIG.SHEET_PAYROLL_BULANAN);
    const employees = getSheetDataAsObjects(CONFIG.SHEET_KARYAWAN);
    const emp = employees.find(e => e.nik === nik);

    if (!emp) return { success: false, message: 'Karyawan tidak ditemukan.' };

    const record = payrollList.find(p => p.nik === nik && p.bulan_tahun === bulanTahun);
    if (!record) {
      return { success: false, message: 'Slip gaji periode ' + bulanTahun + ' belum diterbitkan oleh Admin HRD.' };
    }

    return {
      success: true,
      data: {
        id_payroll: record.id_payroll,
        nik: emp.nik,
        nama: emp.nama,
        departemen: emp.departemen,
        jabatan: emp.jabatan,
        status_karyawan: emp.status_karyawan,
        bank: emp.bank || 'Mandiri',
        no_rekening: emp.no_rekening || '-',
        bulan_tahun: record.bulan_tahun,
        total_hadir: record.total_hadir,
        total_menit_terlambat: record.total_menit_terlambat,
        gaji_pokok: parseFloat(record.gaji_pokok || '0'),
        tunjangan: parseFloat(record.tunjangan || '0'),
        upah_lembur: parseFloat(record.upah_lembur || '0'),
        service_charge: parseFloat(record.service_charge || '0'),
        bpjs_tk: parseFloat(record.bpjs_tk || '0'),
        bpjs_kes: parseFloat(record.bpjs_kes || '0'),
        potongan_pinjaman: parseFloat(record.potongan_pinjaman || '0'),
        denda_terlambat: parseFloat(record.denda_terlambat || '0'),
        thp_bersih: parseFloat(record.thp_bersih || '0'),
        tanggal_transfer: record.tanggal_transfer,
        status_bayar: record.status_bayar
      }
    };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function submitEditProfil(nik, formData) {
  try {
    if (!nik) {
      return { success: false, message: 'Sesi NIK Karyawan tidak terdeteksi. Silakan login ulang.' };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheetEdit = ss.getSheetByName(CONFIG.SHEET_EDIT_PROFIL);
    
    if (!sheetEdit) {
      sheetEdit = ss.insertSheet(CONFIG.SHEET_EDIT_PROFIL);
      sheetEdit.appendRow([
        'id_pengajuan', 'nik', 'nama', 'no_telp', 'no_identitas', 'alamat', 'status_kawin', 'bank', 'no_rekening', 'foto', 'tanggal_pengajuan', 'status_persetujuan'
      ]);
      sheetEdit.getRange('A1:L1').setFontWeight('bold').setBackground('#1E3A8A').setFontColor('#FFFFFF');
    }

    const todayStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm');

    let fotoUrl = formData.current_foto || '';
    if (formData.foto && typeof formData.foto === 'string' && formData.foto.includes('base64,')) {
      fotoUrl = uploadFotoToDrive(formData.foto, nik + '_EditProfil_' + Date.now());
    }

    const newId = 'PRF-' + Date.now().toString().slice(-6);
    sheetEdit.appendRow([
      newId,
      nik,
      formData.nama || '',
      "'" + sanitizePhone(formData.no_telp),
      formData.no_identitas || '',
      formData.alamat || '',
      formData.status_kawin || 'TK/0',
      formData.bank || 'Mandiri',
      formData.no_rekening || '',
      fotoUrl,
      todayStr,
      'Pending'
    ]);

    SpreadsheetApp.flush();
    return { success: true, message: 'Pengajuan perubahan profil berhasil dikirim! Menunggu persetujuan Admin HR.' };
  } catch (err) {
    return { success: false, message: 'Gagal mengajukan edit profil: ' + err.toString() };
  }
}

function approveEditProfil(idPengajuan, actionStatus) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetEdit = ss.getSheetByName(CONFIG.SHEET_EDIT_PROFIL);
    if (!sheetEdit) return { success: false, message: 'Tabel pengajuan edit profil belum tersedia.' };

    const editData = getSheetDataAsObjects(CONFIG.SHEET_EDIT_PROFIL);
    const idx = editData.findIndex(e => e.id_pengajuan === idPengajuan);

    if (idx === -1) return { success: false, message: 'Data pengajuan edit profil tidak ditemukan.' };

    const item = editData[idx];
    sheetEdit.getRange(idx + 2, 12).setValue(actionStatus);

    if (actionStatus === 'Approved') {
      const sheetEmp = ss.getSheetByName(CONFIG.SHEET_KARYAWAN);
      const empData = getSheetDataAsObjects(CONFIG.SHEET_KARYAWAN);
      const empIdx = empData.findIndex(e => e.nik === item.nik);

      if (empIdx !== -1) {
        const rowNum = empIdx + 2;
        if (item.nama) sheetEmp.getRange(rowNum, 2).setValue(item.nama);
        if (item.no_telp) sheetEmp.getRange(rowNum, 3).setValue("'" + sanitizePhone(item.no_telp));
        if (item.no_identitas) sheetEmp.getRange(rowNum, 7).setValue(item.no_identitas);
        if (item.alamat) sheetEmp.getRange(rowNum, 8).setValue(item.alamat);
        if (item.status_kawin) sheetEmp.getRange(rowNum, 9).setValue(item.status_kawin);
        if (item.bank) sheetEmp.getRange(rowNum, 14).setValue(item.bank);
        if (item.no_rekening) sheetEmp.getRange(rowNum, 15).setValue(item.no_rekening);
        if (item.foto && item.foto !== '') sheetEmp.getRange(rowNum, 16).setValue(item.foto);
      }
    }

    SpreadsheetApp.flush();
    return { success: true, message: 'Pengajuan edit profil ' + item.nik + ' berhasil: ' + actionStatus };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function approveKaryawan(nik, actionStatus) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_KARYAWAN);
    const employees = getSheetDataAsObjects(CONFIG.SHEET_KARYAWAN);
    const index = employees.findIndex(e => e.nik === nik);

    if (index === -1) return { success: false, message: 'Data karyawan tidak ditemukan.' };

    const rowNum = index + 2;
    sheet.getRange(rowNum, 19).setValue(actionStatus);

    SpreadsheetApp.flush();
    return { success: true, message: 'Status karyawan ' + nik + ' diperbarui menjadi: ' + actionStatus };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function processAbsensi(nik, userLat, userLong, qrSecretCode, actionType) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetSettings = ss.getSheetByName(CONFIG.SHEET_PENGATURAN);
    const settings = sheetSettings.getDataRange().getDisplayValues()[1] || ['-0.297491', '100.368819', '150', 'BALCONE-QR-2026', '12', '123456', '1000'];
    
    let kantorLatStr = settings[0].toString().trim().replace(',', '.');
    let kantorLongStr = settings[1].toString().trim().replace(',', '.');
    
    if (kantorLatStr === '-297.491') kantorLatStr = '-0.297491';
    if (kantorLongStr === '100.368.819') kantorLongStr = '100.368819';

    const kantorLat = parseFloat(kantorLatStr);
    const kantorLong = parseFloat(kantorLongStr);
    const maxRadius = parseFloat(settings[2]);
    const validQrSecret = settings[3];

    if (qrSecretCode !== validQrSecret) {
      return { success: false, message: 'QR Code Kantor tidak valid!' };
    }

    const distance = haversineDistance(parseFloat(userLat), parseFloat(userLong), kantorLat, kantorLong);
    if (distance > maxRadius) {
      return { 
        success: false, 
        message: 'Posisi Anda di luar radius aman area The Balcone! Jarak Anda: ' + distance + 'm (Maks: ' + maxRadius + 'm)' 
      };
    }

    const todayDate = new Date();
    const todayStr = Utilities.formatDate(todayDate, 'Asia/Jakarta', 'dd/MM/yyyy');
    const isoTodayStr = Utilities.formatDate(todayDate, 'Asia/Jakarta', 'yyyy-MM-dd');
    const timeStr = Utilities.formatDate(todayDate, 'Asia/Jakarta', 'HH:mm:ss');
    
    const rosterList = getSheetDataAsObjects(CONFIG.SHEET_ROSTER);
    const userRosterToday = rosterList.find(r => r.nik === nik && (r.tanggal === isoTodayStr || r.tanggal === todayStr));

    let activeShift = null;
    const shifts = getSheetDataAsObjects(CONFIG.SHEET_SHIFT);

    if (userRosterToday) {
      const shiftVal = userRosterToday.id_shift || userRosterToday.status_hari;
      if (['OFF', 'CT', 'PH', 'EO', 'S', 'I'].includes(shiftVal)) {
        const labelMap = {
          'OFF': 'OFF (Libur)',
          'CT': 'Cuti Tahunan (CT)',
          'PH': 'Publik Holiday (PH)',
          'EO': 'Extra Off (EO)',
          'S': 'Sakit (S)',
          'I': 'Izin (I)'
        };
        return { success: false, message: 'Hari ini jadwal Anda adalah ' + (labelMap[shiftVal] || shiftVal) + ' pada Roster Shift.' };
      }
      activeShift = shifts.find(s => s.id_shift === userRosterToday.id_shift);
    }

    if (!activeShift) {
      activeShift = shifts[0] || { jam_masuk: '08:00', toleransi_terlambat_menit: '15' };
    }

    const sheetAbsensi = ss.getSheetByName(CONFIG.SHEET_ABSENSI);
    const absensiData = getSheetDataAsObjects(CONFIG.SHEET_ABSENSI);
    const existingIndex = absensiData.findIndex(a => a.nik === nik && a.tanggal === todayStr);

    const formattedUserLat = "'" + userLat.toString();
    const formattedUserLong = "'" + userLong.toString();

    if (actionType === 'MASUK') {
      if (existingIndex !== -1 && absensiData[existingIndex].jam_masuk !== '') {
        return { success: false, message: 'Anda sudah melakukan Absen Masuk hari ini!' };
      }

      let lateMinutes = 0;
      let status = 'Tepat Waktu';
      
      const [scheduleHour, scheduleMin] = activeShift.jam_masuk.split(':').map(Number);
      const now = new Date();
      const scheduleTime = new Date(now.getFullYear(), now.getMonth(), now.getDate(), scheduleHour, scheduleMin, 0);
      const toleranceMs = parseInt(activeShift.toleransi_terlambat_menit || '15', 10) * 60 * 1000;
      
      if (now.getTime() > scheduleTime.getTime() + toleranceMs) {
        lateMinutes = Math.round((now.getTime() - scheduleTime.getTime()) / 60000);
        status = 'Terlambat';
      }

      if (existingIndex !== -1) {
        const rowNum = existingIndex + 2;
        sheetAbsensi.getRange(rowNum, 4).setValue(timeStr);
        sheetAbsensi.getRange(rowNum, 5).setNumberFormat('@').setValue(formattedUserLat);
        sheetAbsensi.getRange(rowNum, 6).setNumberFormat('@').setValue(formattedUserLong);
        sheetAbsensi.getRange(rowNum, 10).setValue(status);
        sheetAbsensi.getRange(rowNum, 11).setValue(lateMinutes.toString());
      } else {
        const newId = generateDailyId('ABS', CONFIG.SHEET_ABSENSI);
        sheetAbsensi.appendRow([
          newId, nik, todayStr, timeStr, formattedUserLat, formattedUserLong, '', '', '', status, lateMinutes.toString()
        ]);
      }

      SpreadsheetApp.flush();
      return { success: true, message: 'Absen Masuk Berhasil! Status: ' + status + ' (' + timeStr + ')' };

    } else if (actionType === 'PULANG') {
      if (existingIndex === -1 || absensiData[existingIndex].jam_masuk === '') {
        return { success: false, message: 'Anda belum Absen Masuk hari ini!' };
      }
      if (absensiData[existingIndex].jam_pulang !== '') {
        return { success: false, message: 'Anda sudah Absen Pulang hari ini!' };
      }

      const rowNum = existingIndex + 2;
      sheetAbsensi.getRange(rowNum, 7).setValue(timeStr);
      sheetAbsensi.getRange(rowNum, 8).setNumberFormat('@').setValue(formattedUserLat);
      sheetAbsensi.getRange(rowNum, 9).setNumberFormat('@').setValue(formattedUserLong);

      SpreadsheetApp.flush();
      return { success: true, message: 'Absen Pulang Berhasil! Terima kasih (' + timeStr + ')' };
    }

  } catch (err) {
    return { success: false, message: 'Gagal memproses absensi: ' + err.toString() };
  }
}

function getShiftList() {
  return { success: true, data: getSheetDataAsObjects(CONFIG.SHEET_SHIFT) };
}

function saveShift(idShift, namaShift, jamMasuk, jamPulang, toleransi, oldShiftId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_SHIFT);
    const shifts = getSheetDataAsObjects(CONFIG.SHEET_SHIFT);

    const targetSearchId = oldShiftId || idShift;
    const idx = shifts.findIndex(s => s.id_shift === targetSearchId);

    if (idx !== -1) {
      const row = idx + 2;
      sheet.getRange(row, 1).setValue(idShift);
      sheet.getRange(row, 2).setValue(namaShift);
      sheet.getRange(row, 3).setValue(jamMasuk);
      sheet.getRange(row, 4).setValue(jamPulang);
      sheet.getRange(row, 5).setValue(toleransi.toString());
    } else {
      const newId = idShift || ('SFT-' + (shifts.length + 1).toString().padStart(2, '0'));
      sheet.appendRow([newId, namaShift, jamMasuk, jamPulang, toleransi.toString()]);
    }

    SpreadsheetApp.flush();
    return { success: true, message: 'Shift kerja berhasil disimpan!' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function deleteShift(idShift) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_SHIFT);
    const shifts = getSheetDataAsObjects(CONFIG.SHEET_SHIFT);
    const idx = shifts.findIndex(s => s.id_shift === idShift);

    if (idx !== -1) {
      sheet.deleteRow(idx + 2);
      SpreadsheetApp.flush();
      return { success: true, message: 'Shift berhasil dihapus!' };
    }
    return { success: false, message: 'Shift tidak ditemukan.' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function getDepartemenList() {
  return { success: true, data: getSheetDataAsObjects(CONFIG.SHEET_DEPARTEMEN) };
}

function saveDepartemen(idDept, namaDept) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_DEPARTEMEN);
    const depts = getSheetDataAsObjects(CONFIG.SHEET_DEPARTEMEN);

    if (idDept) {
      const idx = depts.findIndex(d => d.id_departemen === idDept);
      if (idx !== -1) {
        sheet.getRange(idx + 2, 2).setValue(namaDept);
      }
    } else {
      const newId = 'DPT-' + (depts.length + 1).toString().padStart(2, '0');
      sheet.appendRow([newId, namaDept]);
    }

    SpreadsheetApp.flush();
    return { success: true, message: 'Departemen berhasil disimpan!' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function deleteDepartemen(idDept) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_DEPARTEMEN);
    const depts = getSheetDataAsObjects(CONFIG.SHEET_DEPARTEMEN);
    const idx = depts.findIndex(d => d.id_departemen === idDept);

    if (idx !== -1) {
      sheet.deleteRow(idx + 2);
      SpreadsheetApp.flush();
      return { success: true, message: 'Departemen berhasil dihapus!' };
    }
    return { success: false, message: 'Departemen tidak ditemukan.' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function savePengumuman(idPengumuman, judul, isi, status, fotoBase64) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_PENGUMUMAN);
    const list = getSheetDataAsObjects(CONFIG.SHEET_PENGUMUMAN);
    const todayStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy');

    let photoUrl = 'https://images.unsplash.com/photo-1566073771259-6a8506099945?w=600';
    if (fotoBase64 && fotoBase64.includes('base64,')) {
      photoUrl = uploadFotoToDrive(fotoBase64, 'Mading_' + Date.now());
    }

    if (idPengumuman) {
      const idx = list.findIndex(p => p.id_pengumuman === idPengumuman);
      if (idx !== -1) {
        const row = idx + 2;
        sheet.getRange(row, 3).setValue(judul);
        sheet.getRange(row, 4).setValue(isi);
        if (fotoBase64 && fotoBase64.includes('base64,')) {
          sheet.getRange(row, 5).setValue(photoUrl);
        }
        sheet.getRange(row, 7).setValue(status);
      }
    } else {
      const newId = 'PGM-' + Date.now().toString().slice(-5);
      sheet.appendRow([newId, todayStr, judul, isi, photoUrl, 'Admin HR', status || 'Aktif']);
    }

    SpreadsheetApp.flush();
    return { success: true, message: 'Mading pengumuman berhasil disimpan!' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function deletePengumuman(idPengumuman) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_PENGUMUMAN);
    const list = getSheetDataAsObjects(CONFIG.SHEET_PENGUMUMAN);
    const idx = list.findIndex(p => p.id_pengumuman === idPengumuman);

    if (idx !== -1) {
      sheet.deleteRow(idx + 2);
      SpreadsheetApp.flush();
      return { success: true, message: 'Pengumuman Mading berhasil dihapus!' };
    }
    return { success: false, message: 'Pengumuman tidak ditemukan.' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function getKaryawanDashboard(nik, monthYear) {
  try {
    const todayDate = new Date();
    const todayStr = Utilities.formatDate(todayDate, 'Asia/Jakarta', 'dd/MM/yyyy');
    const isoTodayStr = Utilities.formatDate(todayDate, 'Asia/Jakarta', 'yyyy-MM-dd');
    
    const absensi = getSheetDataAsObjects(CONFIG.SHEET_ABSENSI);
    const todayRecord = absensi.find(a => a.nik === nik && a.tanggal === todayStr) || null;
    
    const personalHistory = absensi
      .filter(a => a.nik === nik)
      .reverse()
      .slice(0, 10);

    const announcements = getSheetDataAsObjects(CONFIG.SHEET_PENGUMUMAN)
      .filter(p => p.status_aktif && p.status_aktif.toString().trim().toLowerCase() === 'aktif')
      .reverse();

    const leaveHistory = getSheetDataAsObjects(CONFIG.SHEET_IZIN)
      .filter(i => i.nik === nik)
      .reverse();

    const rosterList = getSheetDataAsObjects(CONFIG.SHEET_ROSTER);
    const shifts = getSheetDataAsObjects(CONFIG.SHEET_SHIFT);

    // Hitung periode cut-off berdasarkan parameter monthYear yang dikirimkan atau otomatis
    let currentMonthStr = monthYear;
    if (!currentMonthStr) {
      let year = todayDate.getFullYear();
      let month = todayDate.getMonth() + 1;
      if (todayDate.getDate() >= 26) {
        month += 1;
        if (month > 12) {
          month = 1;
          year += 1;
        }
      }
      currentMonthStr = year + '-' + String(month).padStart(2, '0');
    }

    const cutOff = getCutoffRange(currentMonthStr);

    // Ambil seluruh roster personal milik karyawan ini dalam siklus cut-off yang dipilih
    const personalRoster = rosterList.filter(r => {
      if (r.nik !== nik || !r.tanggal) return false;
      return r.tanggal >= cutOff.startIsoStr && r.tanggal <= cutOff.endIsoStr;
    }).sort((a, b) => a.tanggal.localeCompare(b.tanggal));

    // Map detail shift untuk tiap tanggal roster personal
    const personalRosterWithDetails = personalRoster.map(r => {
      const shiftVal = r.id_shift || r.status_hari;
      let shiftName = shiftVal;
      let shiftTime = '';

      if (['OFF', 'CT', 'PH', 'EO', 'S', 'I'].includes(shiftVal)) {
        const labelMap = {
          'OFF': 'OFF (Libur)',
          'CT': 'Cuti Tahunan (CT)',
          'PH': 'Publik Holiday (PH)',
          'EO': 'Extra Off (EO)',
          'S': 'Sakit (S)',
          'I': 'Izin (I)'
        };
        shiftName = labelMap[shiftVal] || shiftVal;
      } else {
        const matchedShift = shifts.find(s => s.id_shift === r.id_shift);
        if (matchedShift) {
          shiftName = matchedShift.nama_shift;
          shiftTime = matchedShift.jam_masuk + ' - ' + matchedShift.jam_pulang;
        }
      }

      return {
        id_roster: r.id_roster,
        tanggal: r.tanggal,
        id_shift: r.id_shift,
        status_hari: r.status_hari,
        shift_nama: shiftName,
        shift_jam: shiftTime
      };
    });

    const userRoster = rosterList.find(r => r.nik === nik && (r.tanggal === isoTodayStr || r.tanggal === todayStr));
    
    let shiftInfo = 'Standard / Default';
    if (userRoster) {
      const shiftVal = userRoster.id_shift || userRoster.status_hari;
      if (['OFF', 'CT', 'PH', 'EO', 'S', 'I'].includes(shiftVal)) {
        const labelMap = {
          'OFF': 'OFF (Libur)',
          'CT': 'Cuti Tahunan (CT)',
          'PH': 'Publik Holiday (PH)',
          'EO': 'Extra Off (EO)',
          'S': 'Sakit (S)',
          'I': 'Izin (I)'
        };
        shiftInfo = labelMap[shiftVal] || shiftVal;
      } else {
        const matched = shifts.find(s => s.id_shift === userRoster.id_shift);
        if (matched) shiftInfo = matched.nama_shift + ' (' + matched.jam_masuk + ' - ' + matched.jam_pulang + ')';
        else shiftInfo = userRoster.id_shift;
      }
    }

    return {
      success: true,
      todayRecord: todayRecord,
      history: personalHistory,
      announcements: announcements,
      todayShiftInfo: shiftInfo,
      leaveHistory: leaveHistory,
      personalRoster: personalRosterWithDetails,
      cutoffInfo: {
        startIso: cutOff.startIsoStr,
        endIso: cutOff.endIsoStr
      }
    };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function getHODDashboard(nik, departemen) {
  try {
    const employees = getSheetDataAsObjects(CONFIG.SHEET_KARYAWAN);
    const deptEmployees = employees.filter(e => e.departemen === departemen && e.status_akun === 'Approved' && e.nik !== nik);
    
    const allIzin = getSheetDataAsObjects(CONFIG.SHEET_IZIN);
    
    const pendingIzinHOD = allIzin.filter(i => {
      if (i.status_persetujuan !== 'Pending_HOD' && i.status_persetujuan !== 'Pending') return false;
      const emp = employees.find(e => e.nik === i.nik);
      return emp && emp.departemen === departemen && emp.nik !== nik;
    }).map(i => {
      const emp = employees.find(e => e.nik === i.nik) || {};
      return {
        ...i,
        nama_karyawan: emp.nama || 'N/A'
      };
    });

    const deptIzinHistory = allIzin.filter(i => {
      const emp = employees.find(e => e.nik === i.nik);
      return emp && emp.departemen === departemen;
    }).map(i => {
      const emp = employees.find(e => e.nik === i.nik) || {};
      return {
        ...i,
        nama_karyawan: emp.nama || 'N/A'
      };
    }).reverse();

    return {
      success: true,
      deptEmployees: deptEmployees,
      pendingIzinHOD: pendingIzinHOD,
      deptIzinHistory: deptIzinHistory
    };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function getAdminDashboard() {
  try {
    const todayStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy');
    const absensi = getSheetDataAsObjects(CONFIG.SHEET_ABSENSI);
    const karyawan = getSheetDataAsObjects(CONFIG.SHEET_KARYAWAN);
    const izinList = getSheetDataAsObjects(CONFIG.SHEET_IZIN).reverse();
    const editProfilList = getSheetDataAsObjects(CONFIG.SHEET_EDIT_PROFIL);
    const announcements = getSheetDataAsObjects(CONFIG.SHEET_PENGUMUMAN).reverse();

    const operationalEmployees = karyawan.filter(k => k.role !== 'Admin HR');

    const todayAbsensi = absensi.filter(a => a.tanggal === todayStr);
    const totalHadirToday = todayAbsensi.filter(a => a.status === 'Tepat Waktu' || a.status === 'Terlambat').length;
    const totalTerlambatToday = todayAbsensi.filter(a => a.status === 'Terlambat').length;
    
    const mappedIzinList = izinList.map(i => {
      const emp = karyawan.find(e => e.nik === i.nik) || {};
      return {
        ...i,
        nama: emp.nama || 'N/A',
        departemen: emp.departemen || '-'
      };
    });

    const pendingIzinHRD = mappedIzinList.filter(i => i.status_persetujuan === 'Pending_HRD' || i.status_persetujuan === 'Pending');
    const totalPendingKaryawan = karyawan.filter(k => k.status_akun === 'Pending').length;
    const totalPendingEditProfil = editProfilList.filter(e => e.status_persetujuan === 'Pending').length;

    return {
      success: true,
      stats: {
        totalKaryawan: operationalEmployees.filter(k => k.status_akun === 'Approved' && (k.status_kerja !== 'Resign')).length,
        totalHadirToday: totalHadirToday,
        totalTerlambatToday: totalTerlambatToday,
        totalPendingIzin: pendingIzinHRD.length,
        totalPendingKaryawan: totalPendingKaryawan,
        totalPendingEditProfil: totalPendingEditProfil
      },
      pendingIzin: pendingIzinHRD,
      allIzin: mappedIzinList,
      pendingKaryawan: karyawan.filter(k => k.status_akun === 'Pending'),
      pendingEditProfil: editProfilList.filter(e => e.status_persetujuan === 'Pending'),
      karyawanList: operationalEmployees,
      announcements: announcements
    };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function submitIzinCuti(nik, tglMulai, tglSelesai, jumlahHari, jenis, alasan) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const employees = getSheetDataAsObjects(CONFIG.SHEET_KARYAWAN);
    const emp = employees.find(e => e.nik === nik);

    if (!emp) return { success: false, message: 'Data karyawan tidak ditemukan.' };

    if (jenis === 'Cuti Tahunan') {
      const sisaCuti = parseInt(emp.sisa_cuti_tahunan || '0', 10);
      if (parseInt(jumlahHari, 10) > sisaCuti) {
        return { success: false, message: 'Sisa kuota cuti tahunan Anda (' + sisaCuti + ' hari) tidak mencukupi.' };
      }
    }

    let initialStatus = 'Pending_HOD';
    if (emp.role === 'HOD' || emp.role === 'Admin HR') {
      initialStatus = 'Pending_HRD';
    }

    const sheetIzin = ss.getSheetByName(CONFIG.SHEET_IZIN);
    const newId = 'IZN-' + Date.now().toString().slice(-6);
    sheetIzin.appendRow([
      newId, nik, tglMulai, tglSelesai, jumlahHari, jenis, alasan, initialStatus, ''
    ]);

    SpreadsheetApp.flush();
    const destMessage = initialStatus === 'Pending_HRD' ? 'Admin HRD' : 'HOD Departemen';
    return { success: true, message: 'Pengajuan ' + jenis + ' berhasil dikirim ke ' + destMessage + '.' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function approveIzinCuti(idIzin, action, rejectionReason) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetIzin = ss.getSheetByName(CONFIG.SHEET_IZIN);
    const izinData = getSheetDataAsObjects(CONFIG.SHEET_IZIN);
    const index = izinData.findIndex(i => i.id_izin === idIzin);

    if (index === -1) return { success: false, message: 'Data pengajuan tidak ditemukan.' };

    const item = izinData[index];
    const rowNum = index + 2;
    let newStatus = item.status_persetujuan;

    if (item.status_persetujuan === 'Pending_HOD' || item.status_persetujuan === 'Pending') {
      if (action === 'Approved') {
        newStatus = 'Pending_HRD';
      } else if (action === 'Rejected') {
        newStatus = 'Rejected_HOD';
      }
    } else if (item.status_persetujuan === 'Pending_HRD') {
      if (action === 'Approved') {
        newStatus = 'Approved';
      } else if (action === 'Rejected') {
        newStatus = 'Rejected_HRD';
      }
    }

    sheetIzin.getRange(rowNum, 8).setValue(newStatus);
    if (rejectionReason) {
      sheetIzin.getRange(rowNum, 9).setValue(rejectionReason);
    }

    if (newStatus === 'Approved') {
      if (item.jenis === 'Cuti Tahunan' || item.jenis === 'Cuti') {
        const sheetEmp = ss.getSheetByName(CONFIG.SHEET_KARYAWAN);
        const empData = getSheetDataAsObjects(CONFIG.SHEET_KARYAWAN);
        const empIdx = empData.findIndex(e => e.nik === item.nik);
        
        if (empIdx !== -1) {
          const currentSisa = parseInt(empData[empIdx].sisa_cuti_tahunan || '0', 10);
          const cutiDays = parseInt(item.jumlah_hari, 10);
          sheetEmp.getRange(empIdx + 2, 18).setValue(Math.max(0, currentSisa - cutiDays).toString());
        }
      }

      let shiftCode = 'I';
      if (item.jenis === 'Cuti Tahunan') shiftCode = 'CT';
      else if (item.jenis === 'Publik Holiday') shiftCode = 'PH';
      else if (item.jenis === 'Extra Off') shiftCode = 'EO';
      else if (item.jenis === 'Sakit') shiftCode = 'S';
      else if (item.jenis === 'Izin') shiftCode = 'I';

      let startDate = parseDateStrToDate(item.tanggal_mulai);
      let endDate = parseDateStrToDate(item.tanggal_selesai);

      if (startDate && endDate) {
        const sheetRoster = ss.getSheetByName(CONFIG.SHEET_ROSTER);
        const rosterData = getSheetDataAsObjects(CONFIG.SHEET_ROSTER);

        let cur = new Date(startDate);
        while (cur <= endDate) {
          const isoDate = Utilities.formatDate(cur, 'Asia/Jakarta', 'yyyy-MM-dd');
          
          const rosterIdx = rosterData.findIndex(r => r.nik === item.nik && r.tanggal === isoDate);
          if (rosterIdx !== -1) {
            sheetRoster.getRange(rosterIdx + 2, 4).setValue(shiftCode);
            sheetRoster.getRange(rosterIdx + 2, 5).setValue(shiftCode);
          } else {
            const newId = 'RST-' + Date.now().toString().slice(-6) + '-' + Math.floor(Math.random() * 100);
            sheetRoster.appendRow([newId, item.nik, isoDate, shiftCode, shiftCode]);
          }
          cur.setDate(cur.getDate() + 1);
        }
      }
    }

    SpreadsheetApp.flush();
    return { success: true, message: 'Status pengajuan ' + item.jenis + ' berhasil diperbarui: ' + newStatus };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function calculateMonthlyKPI(bulanTahun) {
  try {
    const cutOff = getCutoffRange(bulanTahun);
    const absensi = getSheetDataAsObjects(CONFIG.SHEET_ABSENSI);
    const karyawan = getSheetDataAsObjects(CONFIG.SHEET_KARYAWAN).filter(k => k.status_akun === 'Approved' && k.role !== 'Admin HR');

    const kpiResults = karyawan.map(emp => {
      const empAbsensi = absensi.filter(a => {
        if (a.nik !== emp.nik || !a.tanggal) return false;
        const parts = a.tanggal.split('/');
        if (parts.length === 3) {
          const recDate = new Date(parseInt(parts[2], 10), parseInt(parts[1], 10) - 1, parseInt(parts[0], 10));
          return recDate >= cutOff.startDate && recDate <= cutOff.endDate;
        }
        return false;
      });

      const totalHadir = empAbsensi.filter(a => a.status === 'Tepat Waktu' || a.status === 'Terlambat').length;
      const totalTerlambat = empAbsensi.filter(a => a.status === 'Terlambat').length;
      
      let totalMenitTerlambat = 0;
      empAbsensi.forEach(a => {
        totalMenitTerlambat += parseInt(a.keterlambatan_menit || '0', 10);
      });

      let skor = 100 - (totalTerlambat * 3) - (totalMenitTerlambat * 0.1);
      skor = Math.max(0, Math.min(100, Math.round(skor)));

      let predikat = 'Sangat Baik';
      if (skor < 70) predikat = 'Perlu Evaluasi';
      else if (skor < 80) predikat = 'Cukup';
      else if (skor < 90) predikat = 'Baik';

      return {
        nik: emp.nik,
        nama: emp.nama,
        departemen: emp.departemen,
        bulan_tahun: bulanTahun + ' (Cut-off 26-25)',
        total_hadir: totalHadir,
        total_terlambat: totalTerlambat,
        total_menit_terlambat: totalMenitTerlambat,
        skor_kpi_persen: skor + '%',
        predikat: predikat
      };
    });

    return { success: true, data: kpiResults };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function getGlobalAttendanceList(searchKey, filterDept, page, limit) {
  try {
    let data = getSheetDataAsObjects(CONFIG.SHEET_ABSENSI);
    const employees = getSheetDataAsObjects(CONFIG.SHEET_KARYAWAN);
    
    data = data.map(item => {
      const emp = employees.find(e => e.nik === item.nik) || {};
      return {
        ...item,
        nama_karyawan: emp.nama || 'N/A',
        departemen: emp.departemen || '-'
      };
    }).reverse();

    if (searchKey) {
      const key = searchKey.toLowerCase();
      data = data.filter(d => 
        d.nama_karyawan.toLowerCase().includes(key) || 
        d.tanggal.includes(key) || 
        d.nik.toLowerCase().includes(key)
      );
    }

    if (filterDept && filterDept !== 'ALL') {
      data = data.filter(d => d.departemen === filterDept);
    }

    const totalRecords = data.length;
    const startIndex = (page - 1) * limit;
    const paginatedData = data.slice(startIndex, startIndex + limit);

    return {
      success: true,
      data: paginatedData,
      total: totalRecords,
      totalPages: Math.ceil(totalRecords / limit),
      currentPage: page
    };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function generatePDFReport(bulanTahun) {
  try {
    const kpiRes = calculateMonthlyKPI(bulanTahun);
    if (!kpiRes.success) return kpiRes;

    let htmlContent = '<h1 style="text-align:center; color:#1E3A8A;">THE BALCONE SUITES & RESORT</h1>';
    htmlContent += '<h3 style="text-align:center; color:#555;">LAPORAN KEHADIRAN & KPI KARYAWAN PERIODE CUT-OFF (' + bulanTahun + ')</h3>';
    htmlContent += '<p style="text-align:center; font-size:12px; color:#777;">Siklus Cut-Off Resmi: Tanggal 26 s/d Tanggal 25</p>';
    htmlContent += '<table border="1" style="width:100%; border-collapse:collapse; font-family:sans-serif; margin-top:15px;" cellpadding="6">';
    htmlContent += '<tr style="background-color:#1E3A8A; color:white;">' +
      '<th>NIK</th><th>Nama Karyawan</th><th>Departemen</th><th>Total Hadir</th><th>Terlambat</th><th>Skor KPI</th><th>Predikat</th>' +
      '</tr>';

    kpiRes.data.forEach(row => {
      htmlContent += '<tr>' +
        '<td>' + row.nik + '</td>' +
        '<td>' + row.nama + '</td>' +
        '<td>' + row.departemen + '</td>' +
        '<td style="text-align:center;">' + row.total_hadir + '</td>' +
        '<td style="text-align:center;">' + row.total_terlambat + '</td>' +
        '<td style="text-align:center; font-weight:bold;">' + row.skor_kpi_persen + '</td>' +
        '<td style="text-align:center;">' + row.predikat + '</td>' +
        '</tr>';
    });
    htmlContent += '</table>';

    const blob = Utilities.newBlob(htmlContent, 'text/html', 'Laporan_Balcone_' + bulanTahun.replace('/', '_') + '.html');
    const pdfBlob = blob.getAs('application/pdf');
    const base64Pdf = Utilities.base64Encode(pdfBlob.getBytes());

    return {
      success: true,
      pdfBase64: 'data:application/pdf;base64,' + base64Pdf,
      fileName: 'Laporan_Kehadiran_Balcone_' + bulanTahun.replace('/', '_') + '.pdf'
    };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function getOfficeSettings() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_PENGATURAN);
    const settings = sheet.getDataRange().getDisplayValues()[1] || ['-0.297491', '100.368819', '150', 'BALCONE-QR-2026', '12', '123456', '1000'];
    
    let lat = settings[0].toString().trim().replace(',', '.');
    let long = settings[1].toString().trim().replace(',', '.');

    if (lat === '-297.491') lat = '-0.297491';
    if (long === '100.368.819') long = '100.368819';

    return {
      success: true,
      settings: {
        lat: lat,
        long: long,
        radius: settings[2],
        qrCode: settings[3],
        pinPayroll: settings[5] || '123456',
        rateDendaDefault: settings[6] || '1000'
      }
    };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function saveOfficeSettings(lat, long, radius, qrCode, pinPayroll, rateDenda) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_PENGATURAN);
    
    sheet.getRange('A2:B2').setNumberFormat('@');
    
    let cleanLat = lat.toString().trim().replace(',', '.');
    let cleanLong = long.toString().trim().replace(',', '.');
    
    sheet.getRange(2, 1).setValue("'" + cleanLat);
    sheet.getRange(2, 2).setValue("'" + cleanLong);
    sheet.getRange(2, 3).setValue(radius.toString());
    sheet.getRange(2, 4).setValue(qrCode);
    if (pinPayroll) sheet.getRange(2, 6).setValue(pinPayroll.toString());
    if (rateDenda) sheet.getRange(2, 7).setValue(rateDenda.toString());
    
    SpreadsheetApp.flush();
    return { success: true, message: 'Pengaturan area kantor, PIN Payroll, & Denda berhasil diperbarui!' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function getRosterData(monthYear, filterDept) {
  try {
    const cutOff = getCutoffRange(monthYear);
    let karyawan = getSheetDataAsObjects(CONFIG.SHEET_KARYAWAN).filter(k => k.status_akun === 'Approved' && k.role !== 'Admin HR');
    
    karyawan = karyawan.filter(emp => {
      if (!emp.tgl_resign || emp.tgl_resign === '') return true;
      return emp.tgl_resign >= cutOff.startIsoStr;
    });

    if (filterDept && filterDept !== 'ALL') {
      karyawan = karyawan.filter(k => k.departemen === filterDept);
    }

    const rosterList = getSheetDataAsObjects(CONFIG.SHEET_ROSTER);
    const shifts = getSheetDataAsObjects(CONFIG.SHEET_SHIFT);

    const filteredRoster = rosterList.filter(r => {
      if (!r.tanggal) return false;
      return r.tanggal >= cutOff.startIsoStr && r.tanggal <= cutOff.endIsoStr;
    });

    return {
      success: true,
      karyawan: karyawan,
      roster: filteredRoster,
      shifts: shifts,
      cutoff: {
        startIso: cutOff.startIsoStr,
        endIso: cutOff.endIsoStr
      }
    };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function saveBulkRoster(rosterItems) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetRoster = ss.getSheetByName(CONFIG.SHEET_ROSTER);
    const existingRoster = getSheetDataAsObjects(CONFIG.SHEET_ROSTER);

    rosterItems.forEach(item => {
      const idx = existingRoster.findIndex(r => r.nik === item.nik && r.tanggal === item.tanggal);
      if (idx !== -1) {
        const rowNum = idx + 2;
        sheetRoster.getRange(rowNum, 4).setValue(item.id_shift);
        sheetRoster.getRange(rowNum, 5).setValue(item.status_hari || item.id_shift);
      } else {
        const newId = 'RST-' + Date.now().toString().slice(-6) + '-' + Math.floor(Math.random() * 100);
        sheetRoster.appendRow([
          newId, item.nik, item.tanggal, item.id_shift, item.status_hari || item.id_shift
        ]);
      }
    });

    SpreadsheetApp.flush();
    return { success: true, message: 'Jadwal Roster Shift berhasil disimpan!' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}