/**
 * FILE: 10_ControlAbsensi.gs
 * STATUS: TARGET TERKUNCI + ANTI DATA GANDA + DETEKSI MANGKIR (TANPA KETERANGAN)
 */
function getLapControlAbsensi(tglMulai, tglSelesai) {
  try {
    // =========================================================================
    // 1. PASTE ID SPREADSHEET "Database HRIS KOBE" BAPAK DI SINI
    // =========================================================================
    const idSpreadsheet = "1VOMGDOXD7C1iONXgdpOzxyc4Xmzueow-yQHsAwPQfUw"; 
    
    const ss = SpreadsheetApp.openById(idSpreadsheet);
    const sheetAbs = ss.getSheetByName("Raw_Kehadiran"); 
    const sheetMst = ss.getSheetByName("Master_Karyawan"); // WAJIB ADA UNTUK BUKU INDUK

    if (!sheetAbs) throw new Error("Sheet 'Raw_Kehadiran' tidak ditemukan!");
    if (!sheetMst) throw new Error("Sheet 'Master_Karyawan' tidak ditemukan!");

    const dataAbs = sheetAbs.getDataRange().getValues();
    const dataMst = sheetMst.getDataRange().getValues();

    // 2. Ambil Daftar Seluruh Karyawan Aktif (Asumsi: NRPP di Kolom A, Nama di Kolom B)
    let daftarKaryawan = [];
    for (let i = 1; i < dataMst.length; i++) {
      let nrppMst = dataMst[i][0];
      let namaMst = dataMst[i][1];
      if (nrppMst && namaMst) {
        daftarKaryawan.push({ nrpp: nrppMst, nama: namaMst });
      }
    }

    // Parsing Rentang Tanggal
    let partsStart = tglMulai.split("-");
    let start = new Date(partsStart[0], partsStart[1] - 1, partsStart[2]);
    start.setHours(0, 0, 0, 0);

    let partsEnd = tglSelesai.split("-");
    let end = new Date(partsEnd[0], partsEnd[1] - 1, partsEnd[2]);
    end.setHours(23, 59, 59, 999);
    
    // Mesin Pencatat Absensi Harian
    let recordKehadiran = {}; // Menyimpan status: recordKehadiran["TANGGAL_NRPP"] = "Status"
    let daftarTanggalKerja = new Set(); // Menyimpan tanggal berapa saja yang merupakan hari kerja aktif

    // 3. BACA RAW DATA: Kumpulkan siapa saja yang hadir/ada keterangan
    for (let i = 1; i < dataAbs.length; i++) {
      let tglValue = dataAbs[i][0];    // Kolom A: Tanggal
      let nrppRaw = dataAbs[i][1];     // Kolom B: NRPP
      let statusValue = dataAbs[i][5]; // Kolom F: Status

      if (!tglValue || !nrppRaw) continue;

      let tgl;
      if (tglValue instanceof Date) {
        tgl = tglValue;
      } else {
        let parts = String(tglValue).split("/");
        if(parts.length === 3) tgl = new Date(parts[2], parts[1] - 1, parts[0]); 
        else tgl = new Date(tglValue); 
      }
      
      // Jika masuk rentang kalender
      if (!isNaN(tgl.getTime()) && tgl.getTime() >= start.getTime() && tgl.getTime() <= end.getTime()) {
        let dd = String(tgl.getDate()).padStart(2, '0');
        let mm = String(tgl.getMonth() + 1).padStart(2, '0');
        let yyyy = tgl.getFullYear();
        let tanggalStr = `${dd}/${mm}/${yyyy}`;
        
        // Catat bahwa tanggal ini adalah hari kerja (ada aktivitas absen)
        daftarTanggalKerja.add(tanggalStr);
        
        let status = String(statusValue || "").toUpperCase().trim();
        let kunciAbsen = tanggalStr + "_" + nrppRaw;
        
        // Jika statusnya kosong di Excel, kita anggap dia "HADIR" (karena dia nge-tap jari)
        if (status === "") status = "HADIR";
        
        // Simpan statusnya
        recordKehadiran[kunciAbsen] = status;
      }
    }

    let hasil = [];

    // 4. LOGIKA TOP-DOWN: Bandingkan Buku Induk dengan Record Absensi
    // Kita cek satu-satu setiap HARI KERJA
    daftarTanggalKerja.forEach(function(tanggalKerja) {
      
      // Kita absen satu-satu setiap KARYAWAN di hari tersebut
      for (let j = 0; j < daftarKaryawan.length; j++) {
        let emp = daftarKaryawan[j];
        let kunciCek = tanggalKerja + "_" + emp.nrpp;
        
        let statusKaryawanHariItu = recordKehadiran[kunciCek];
        
        if (!statusKaryawanHariItu) {
          // KONDISI 1: TIDAK ADA DATA SAMA SEKALI DI RAW_KEHADIRAN (SI D MANGKIR)
          hasil.push({
            nama: emp.nama,
            tanggal: tanggalKerja,
            keterangan: "TANPA KETERANGAN"
          });
        } else {
          // KONDISI 2: ADA DATA, TAPI STATUSNYA SAKIT/CUTI/PG/ALPA (SI A, B, C)
          if (statusKaryawanHariItu === "S" || statusKaryawanHariItu === "SAKIT" || 
              statusKaryawanHariItu === "C" || statusKaryawanHariItu === "CUTI" || 
              statusKaryawanHariItu === "PG" || statusKaryawanHariItu === "ALPA") {
            hasil.push({
              nama: emp.nama,
              tanggal: tanggalKerja,
              keterangan: statusKaryawanHariItu
            });
          }
        }
      }
    });

    // Urutkan berdasarkan Tanggal lalu Nama (opsional agar rapi)
    hasil.sort((a, b) => a.tanggal.localeCompare(b.tanggal) || a.nama.localeCompare(b.nama));

    return hasil;
  } catch (error) {
    throw new Error(error.message); 
  }
}
