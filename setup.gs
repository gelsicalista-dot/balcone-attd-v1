/**
 * THE BALCONE SUITES & RESORT - Database Migration & Setup Module
 * File: setup.gs
 * Author: ZettBOT Assistant by Zettbos
 */

const SHEETS = {
  KARYAWAN: 'Tabel_Karyawan',
  SHIFT: 'Tabel_Shift',
  ABSENSI: 'Tabel_Absensi',
  IZIN_CUTI: 'Tabel_Izin_Cuti',
  KPI: 'Tabel_KPI',
  PENGUMUMAN: 'Tabel_Pengumuman',
  PENGATURAN: 'Tabel_Pengaturan',
  DEPARTEMEN: 'Tabel_Departemen',
  ROSTER: 'Tabel_Roster',
  EDIT_PROFIL: 'Tabel_Edit_Profil',
  GAJI_MASTER: 'Tabel_Gaji_Master',
  PAYROLL_BULANAN: 'Tabel_Payroll_Bulanan',
  PINJAMAN: 'Tabel_Pinjaman_Karyawan'
};

/**
 * FUNGSI KHUSUS PEMAKSA OTORISASI GOOGLE DRIVE
 */
function authorizeDrive() {
  const folderName = 'Balcone_Foto_Karyawan';
  const folders = DriveApp.getFoldersByName(folderName);
  let folder;
  if (folders.hasNext()) {
    folder = folders.next();
  } else {
    folder = DriveApp.createFolder(folderName);
  }
  folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  Logger.log('OTORISASI BERHASIL! Folder Google Drive "' + folderName + '" siap digunakan.');
}

/**
 * Inisialisasi & Migrasi Database (Safe Migrate + Force Header Sync)
 */
function setupDatabase() {
  const folderName = 'Balcone_Foto_Karyawan';
  const folders = DriveApp.getFoldersByName(folderName);
  let folder;
  if (!folders.hasNext()) {
    folder = DriveApp.createFolder(folderName);
  } else {
    folder = folders.next();
  }
  folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Skema Header Kolom Resmi THE BALCONE SUITES & RESORT
  const schema = {
    [SHEETS.KARYAWAN]: [
      'nik', 'nama', 'no_telp', 'password', 'role', 
      'status_karyawan', 'no_identitas', 'alamat', 'status_kawin', 
      'departemen', 'jabatan', 'tgl_masuk', 'tgl_lahir', 'bank', 'no_rekening', 
      'foto', 'kuota_cuti_tahunan', 'sisa_cuti_tahunan', 'status_akun', 'status_kerja', 'tgl_resign'
    ],
    [SHEETS.SHIFT]: [
      'id_shift', 'nama_shift', 'jam_masuk', 'jam_pulang', 'toleransi_terlambat_menit'
    ],
    [SHEETS.DEPARTEMEN]: [
      'id_departemen', 'nama_departemen'
    ],
    [SHEETS.ABSENSI]: [
      'id_absen', 'nik', 'tanggal', 'jam_masuk', 'lat_masuk', 
      'long_masuk', 'jam_pulang', 'lat_pulang', 'long_pulang', 'status', 'keterlambatan_menit'
    ],
    [SHEETS.IZIN_CUTI]: [
      'id_izin', 'nik', 'tanggal_mulai', 'tanggal_selesai', 
      'jumlah_hari', 'jenis', 'alasan', 'status_persetujuan', 'alasan_penolakan'
    ],
    [SHEETS.KPI]: [
      'id_kpi', 'nik', 'bulan_tahun', 'total_hadir', 
      'total_terlambat', 'total_menit_terlambat', 'total_izin', 'skor_kpi_persen', 'predikat'
    ],
    [SHEETS.PENGUMUMAN]: [
      'id_pengumuman', 'tanggal_post', 'judul', 'isi_pengumuman', 'foto_pengumuman', 'pembuat', 'status_aktif'
    ],
    [SHEETS.PENGATURAN]: [
      'koordinat_kantor_lat', 'koordinat_kantor_long', 'radius_meter', 'qr_secret_code', 'default_kuota_cuti', 'pin_payroll', 'rate_denda_per_menit_default'
    ],
    [SHEETS.ROSTER]: [
      'id_roster', 'nik', 'tanggal', 'id_shift', 'status_hari'
    ],
    [SHEETS.EDIT_PROFIL]: [
      'id_pengajuan', 'nik', 'nama', 'no_telp', 'no_identitas', 'alamat', 'status_kawin', 'bank', 'no_rekening', 'foto', 'tanggal_pengajuan', 'status_persetujuan'
    ],
    [SHEETS.GAJI_MASTER]: [
      'nik', 'gaji_pokok', 'tunjangan_jabatan', 'tunjangan_makan', 'tunjangan_transport', 'rate_denda_per_menit', 'bpjs_tk_aktif', 'bpjs_kes_aktif', 'mode_prorate'
    ],
    [SHEETS.PAYROLL_BULANAN]: [
      'id_payroll', 'nik', 'bulan_tahun', 'total_hadir', 'total_menit_terlambat', 'gaji_pokok', 'tunjangan', 'upah_lembur', 'service_charge', 'bpjs_tk', 'bpjs_kes', 'potongan_pinjaman', 'denda_terlambat', 'thp_bersih', 'tanggal_transfer', 'status_bayar'
    ],
    [SHEETS.PINJAMAN]: [
      'id_pinjaman', 'nik', 'tanggal_pinjam', 'total_pinjaman', 'cicilan_per_bulan', 'sisa_pinjaman', 'status_lunas'
    ]
  };

  Object.keys(schema).forEach(sheetName => {
    let sheet = ss.getSheetByName(sheetName);
    const expectedHeaders = schema[sheetName];

    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow(expectedHeaders);
    } else {
      // Safe Migrate: Update header row 1 tanpa merusak isi data
      sheet.getRange(1, 1, 1, expectedHeaders.length).setValues([expectedHeaders]);
    }

    const headerRange = sheet.getRange(1, 1, 1, expectedHeaders.length);
    headerRange.setFontWeight('bold')
               .setBackground('#1E3A8A')
               .setFontColor('#FFFFFF');
    sheet.setFrozenRows(1);
  });

  seedInitialData(ss);

  // Auto-Correction format koordinat
  const sheetPengaturan = ss.getSheetByName(SHEETS.PENGATURAN);
  if (sheetPengaturan) {
    sheetPengaturan.getRange('A2:B2').setNumberFormat('@');
    const data = sheetPengaturan.getDataRange().getDisplayValues();
    if (data.length > 1) {
      let latVal = data[1][0];
      let longVal = data[1][1];

      if (latVal === '-297.491' || latVal === '-297,491') latVal = '-0.297491';
      if (longVal === '100.368.819' || longVal === '100,368819') longVal = '100.368819';

      sheetPengaturan.getRange(2, 1).setValue("'" + latVal);
      sheetPengaturan.getRange(2, 2).setValue("'" + longVal);
    }
  }

  SpreadsheetApp.flush();
  Logger.log('Inisialisasi & Safe Migrate Database Balcone Resort Berhasil!');
}

function seedInitialData(ss) {
  // 1. Seed Tabel Pengaturan Kantor
  const sheetPengaturan = ss.getSheetByName(SHEETS.PENGATURAN);
  sheetPengaturan.getRange('A2:B2').setNumberFormat('@');
  if (sheetPengaturan.getLastRow() <= 1) {
    sheetPengaturan.appendRow(["'-0.297491", "'100.368819", "150", "BALCONE-QR-2026", "12", "123456", "1000"]);
  }

  // 2. Seed Tabel Shift Kerja
  const sheetShift = ss.getSheetByName(SHEETS.SHIFT);
  if (sheetShift.getLastRow() <= 1) {
    sheetShift.appendRow(['SFT-01', 'Shift Pagi', '08:00', '17:00', '15']);
    sheetShift.appendRow(['SFT-02', 'Shift Middle', '12:00', '21:00', '15']);
    sheetShift.appendRow(['SFT-03', 'Shift Malam', '22:00', '07:00', '10']);
  }

  // 3. Seed Tabel Departemen
  const sheetDept = ss.getSheetByName(SHEETS.DEPARTEMEN);
  if (sheetDept.getLastRow() <= 1) {
    const initialDepts = [
      ['DPT-01', 'Front Office'],
      ['DPT-02', 'Housekeeping'],
      ['DPT-03', 'Engineering'],
      ['DPT-04', 'FB Service'],
      ['DPT-05', 'FB Product'],
      ['DPT-06', 'HRD Dept'],
      ['DPT-07', 'Accounting'],
      ['DPT-08', 'Sales & Marketing']
    ];
    initialDepts.forEach(d => sheetDept.appendRow(d));
  }

  // 4. Seed Tabel Karyawan & Master Gaji Default
  const sheetKaryawan = ss.getSheetByName(SHEETS.KARYAWAN);
  const updatedData = sheetKaryawan.getDataRange().getDisplayValues();
  let hasAdmin = false;
  for (let i = 1; i < updatedData.length; i++) {
    const role = updatedData[i][4];
    const userTelp = updatedData[i][2];
    if (role === 'Admin HR' || userTelp.toLowerCase() === 'admin') {
      hasAdmin = true;
      break;
    }
  }

  if (!hasAdmin) {
    sheetKaryawan.appendRow([
      'NIK-0000', 'Admin HR Balcone', 'Admin', 'Admin', 'Admin HR', 
      'PKWTT', '1234567890123456', 'Bukittinggi, Sumatra Barat', 'K/1', 
      'HRD Dept', 'HR Manager', '2022-01-01', '1990-01-01', 'BCA', '123456789', 
      'https://images.unsplash.com/photo-1573496359142-b8d87734a5a2?w=150', '12', '12', 'Approved', 'Aktif', ''
    ]);
  }

  if (sheetKaryawan.getLastRow() <= 1) {
    sheetKaryawan.appendRow([
      'NIK-0002', 'Budi Santoso', "'081987654321", '123456', 'Karyawan', 
      'PKWT', '3201234567890001', 'Jl. Raya Bukittinggi No. 45', 'TK/0', 
      'Front Office', 'Receptionist', '2023-05-10', '1998-05-15', 'Mandiri', '987654321', 
      'https://images.unsplash.com/photo-1535713875002-d1d0cf377fde?w=150', '12', '10', 'Approved', 'Aktif', ''
    ]);
  }

  // Seed Master Gaji Karyawan Default jika belum ada
  const sheetGajiMaster = ss.getSheetByName(SHEETS.GAJI_MASTER);
  if (sheetGajiMaster.getLastRow() <= 1) {
    sheetGajiMaster.appendRow(['NIK-0002', '3500000', '500000', '300000', '200000', '1000', 'Ya', 'Ya', 'Otomatis']);
  }

  // 5. Seed Tabel Pengumuman Mading
  const sheetPengumuman = ss.getSheetByName(SHEETS.PENGUMUMAN);
  if (sheetPengumuman.getLastRow() <= 1) {
    const todayStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy');
    sheetPengumuman.appendRow([
      'PGM-001', todayStr, 'Selamat Datang di Portal Mading Digital Balcone Resort', 
      'Seluruh karyawan diwajibkan melakukan scan QR Code Pass & Verifikasi GPS saat jam masuk dan pulang kerja.', 
      'https://images.unsplash.com/photo-1566073771259-6a8506099945?w=600',
      'Admin HR', 'Aktif'
    ]);
  }
}