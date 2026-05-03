function getDbInstance() {
  return SpreadsheetApp.openById(appConfig().DB_UTAMA);
}

function setupDatabase() {
  const ss = getDbInstance();
  const sheets = [
    { name: 'Master_Karyawan', header: ['NRPP', 'Nama Karyawan', 'Gol', 'Status', 'Jabatan', 'dept', 'Divisi', 'Perusahaan'], color: '#d9ead3' },
    { name: 'Raw_Kehadiran', header: ['Tanggal', 'NRPP', 'Nama', 'Jam Masuk', 'Jam Keluar', 'Status'], color: '#c9daf8' },
    { name: 'Rekap_Bulanan', header: ['Bulan-Tahun', 'NRPP', 'Nama', 'Divisi', 'H (Hadir)', 'S (Sakit)', 'C (Cuti)', 'PG (Potong Gaji/Alpa)', 'ST', 'Visit Customer'], color: '#fff2cc' },
    { name: 'Master_Libur', header: ['Tanggal', 'Status', 'Keterangan'], color: '#f4cccc' },
    // TABLE BARU: MASTER USER
    { name: 'Master_User', header: ['Username', 'Password', 'Nama Lengkap', 'Role'], color: '#cfe2f3' }
  ];

  sheets.forEach(s => {
    if (!ss.getSheetByName(s.name)) {
      let sheet = ss.insertSheet(s.name);
      sheet.appendRow(s.header);
      sheet.getRange(1, 1, 1, s.header.length).setFontWeight("bold").setBackground(s.color);
      
      // Khusus Master_User, tambahkan 1 akun default: admin / admin123
      if (s.name === 'Master_User') {
        sheet.appendRow(['admin', 'admin123', 'Administrator Utama', 'SuperAdmin']);
      }
    }
  });
}
