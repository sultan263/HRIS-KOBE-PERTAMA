function getCrosstabRekapData(startStr, endStr, perusahaan) {
  const ss = getDbInstance();
  const datesArr = [];
  const namaHari = ["Minggu","Senin","Selasa","Rabu","Kamis","Jumat","Sabtu"];
  const namaBulan = ["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"];

  const liburSheet = ss.getSheetByName('Master_Libur');
  let liburMap = {};
  if(liburSheet) {
    let lData = liburSheet.getDataRange().getDisplayValues();
    for(let i=1; i<lData.length; i++) if(lData[i][0]) liburMap[lData[i][0].toString().trim()] = { stat: lData[i][1], ket: lData[i][2] };
  }

  let totalWorkdays = 0;
  for (let d = new Date(startStr); d <= new Date(endStr); d.setDate(d.getDate() + 1)) {
    let raw = Utilities.formatDate(d, "GMT+7", "dd/MM/yyyy");
    let isWeekend = (d.getDay() === 0 || d.getDay() === 6);
    let libur = liburMap[raw];
    
    datesArr.push({ 
      raw: raw, display: `${namaHari[d.getDay()]}, ${Utilities.formatDate(d, "GMT+7", "dd")} ${namaBulan[d.getMonth()]} ${d.getFullYear()}`,
      isLibur: !!libur, liburStat: libur ? libur.stat : '', liburKet: libur ? libur.ket : ''
    });
    if(!isWeekend && !libur) totalWorkdays++;
  }

  const masterData = ss.getSheetByName('Master_Karyawan').getDataRange().getDisplayValues();
  let employeesMap = {}; let employees = [];
  for(let i=1; i<masterData.length; i++) {
    let row = masterData[i];
    if(!row || !row[0]) continue;
    let pt = (row[7] || "").toString().trim();
    if(perusahaan === "ALL" || pt === perusahaan) {
      let nrpp = row[0].toString().trim();
      let emp = { nrpp: nrpp, nama: row[1] || "-", H:0, S:0, C:0, ST:0, VC:0, PG:0, LN:0, CB:0, lateSec: 0 };
      employees.push(emp); employeesMap[nrpp] = emp;
    }
  }

  const rawKehadiran = ss.getSheetByName('Raw_Kehadiran').getDataRange().getDisplayValues();
  let attMap = {}; 
  for(let i=1; i<rawKehadiran.length; i++) {
    let row = rawKehadiran[i]; 
    if(!row || !row[0] || !row[1]) continue;
    let tgl = row[0].toString().trim(); 
    let nrpp = row[1].toString().trim();

    if(!attMap[nrpp]) attMap[nrpp] = {};
    let status = (row[5] || "H").toString().trim().toUpperCase();
    attMap[nrpp][tgl] = { in: row[3]||"-", out: row[4]||"-", stat: status };

    if (isDateInRange(tgl, startStr, endStr) && employeesMap[nrpp]) {
      let emp = employeesMap[nrpp];
      if(status==='H') emp.H++; else if(status==='S') emp.S++; else if(status==='C') emp.C++;
      else if(status==='ST') emp.ST++; else if(status==='PG') emp.PG++; else if(status==='VC') emp.VC++;
      else if(status==='LN') emp.LN++; else if(status==='CB') emp.CB++;

      let masuk = row[3];
      if (masuk) {
        let stdStart = timeToSeconds(getWorkingHours(tgl).START);
        let mSec = timeToSeconds(masuk);
        if (mSec > stdStart) emp.lateSec += (mSec - stdStart);
      }
    }
  }
  employees.forEach(e => e.lateStr = secondsToTimeStr(e.lateSec));
  return { dates: datesArr, employees: employees, attendance: attMap, workdays: totalWorkdays };
}

function getMultiRekapSummary(nrppArray, startDateStr, endDateStr) {
  const ss = getDbInstance();
  const raw = ss.getSheetByName('Raw_Kehadiran').getDataRange().getDisplayValues();
  const master = ss.getSheetByName('Master_Karyawan').getDataRange().getDisplayValues();
  
  let mLookup = {};
  for(let i=1; i<master.length; i++) {
    if(master[i] && master[i][0]) mLookup[master[i][0].toString().trim()] = { nama: master[i][1]||"-", div: master[i][6]||"-" };
  }
  
  let result = {};
  for(let i=1; i<raw.length; i++) {
    let row = raw[i];
    if(!row || !row[0] || !row[1]) continue;
    let tgl = row[0].toString().trim(); 
    let nrpp = row[1].toString().trim(); 

    if (nrppArray.includes(nrpp) && isDateInRange(tgl, startDateStr, endDateStr)) {
      if(!result[nrpp]) {
        let m = mLookup[nrpp] || { nama: "-", div: "-" };
        result[nrpp] = { nrpp: nrpp, nama: m.nama, div: m.div, H:0, S:0, C:0, PG:0, totalLate: 0 };
      }
      let status = (row[5]||"H").toString().trim().toUpperCase();
      if(status==='H') result[nrpp].H++; else if(status==='S') result[nrpp].S++; else if(status==='C') result[nrpp].C++; else if(status==='PG') result[nrpp].PG++;
      
      let masuk = row[3];
      if(masuk) {
        let stdStartSec = timeToSeconds(getWorkingHours(tgl).START);
        let mSec = timeToSeconds(masuk);
        if(mSec > stdStartSec) result[nrpp].totalLate += (mSec - stdStartSec);
      }
    }
  }
  
  return Object.values(result).map(r => [ `${startDateStr} s/d ${endDateStr}`, r.nrpp, r.nama, r.div, r.H, r.S, r.C, r.PG, secondsToTimeStr(r.totalLate) ]);
}

function getLaporanIndividuDynamicByDate(nrpp, startStr, endStr) {
  const ss = getDbInstance();
  const masterData = ss.getSheetByName('Master_Karyawan').getDataRange().getDisplayValues();
  
  // RESTORE PENTING: Nilai Default agar tidak blank
  let profil = { nrpp: nrpp, nama: "-", gol: "-", status: "-", jabatan: "-", dept: "-", divisi: "-", pt: "" };
  for(let i=1; i<masterData.length; i++) {
    let row = masterData[i];
    if(!row || !row[0]) continue;
    // RESTORE PENTING: .trim() agar pencarian presisi
    if(row[0].toString().trim() === nrpp.toString().trim()) {
      profil.nama = row[1] || "-"; profil.gol = row[2] || "-"; profil.status = row[3] || "-"; profil.jabatan = row[4] || "-";
      profil.dept = row[5] || "-"; profil.divisi = row[6] || "-"; profil.pt = row[7] || ""; break;
    }
  }

  const liburSheet = ss.getSheetByName('Master_Libur');
  let liburMap = {};
  if(liburSheet) {
    let lData = liburSheet.getDataRange().getDisplayValues();
    for(let i=1; i<lData.length; i++) if(lData[i][0]) liburMap[lData[i][0].toString().trim()] = { stat: lData[i][1], ket: lData[i][2] };
  }

  const sheetRaw = ss.getSheetByName('Raw_Kehadiran');
  let history = [];
  const namaHariList = ["Minggu", "Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu"];

  if (sheetRaw) {
    const rawData = sheetRaw.getDataRange().getDisplayValues();
    for(let i=1; i<rawData.length; i++) {
      let row = rawData[i];
      if (!row || !row[0] || !row[1]) continue;
      if(row[1].toString().trim() !== nrpp.toString().trim()) continue;
      if (!isDateInRange(row[0], startStr, endStr)) continue;

      let tgl = row[0].toString().trim();
      let masukTeks = (row[3] || "").toString().trim(); 
      let keluarTeks = (row[4] || "").toString().trim();
      let stat = (row[5] || "H").toString().trim().toUpperCase();

      let parts = tgl.split('/');
      if(parts.length !== 3) continue;
      let dObj = new Date(parseInt(parts[2]), parseInt(parts[1])-1, parseInt(parts[0]));
      let namaHari = namaHariList[dObj.getDay()];

      let stdStartSec = timeToSeconds(getWorkingHours(tgl).START);
      let mSec = timeToSeconds(masukTeks);
      let kSec = timeToSeconds(keluarTeks);
      
      let strTerlambat = "-"; let strDurasi = "-";
      if (masukTeks && mSec > 0) {
          let terlambatSec = (mSec > stdStartSec) ? (mSec - stdStartSec) : 0;
          strTerlambat = secondsToTimeStr(terlambatSec); 
      }
      if (masukTeks && keluarTeks && kSec > 0 && mSec > 0) {
          let durasiSec = (kSec > mSec) ? (kSec - mSec) : 0;
          strDurasi = secondsToTimeStr(durasiSec);
      }

      history.push({
        hari: namaHari, tgl: tgl, masuk: formatTimeForInput(masukTeks) || '-', keluar: formatTimeForInput(keluarTeks) || '-',
        terlambat: strTerlambat, durasi: strDurasi, status: stat
      });
    }
  }
  
  let summary = { H: 0, S: 0, C: 0, ST: 0, PG: 0, VC: 0, TotalHariKerja: 0 };
  summary.H = history.filter(r => r.status === 'H').length;
  summary.S = history.filter(r => r.status === 'S').length;
  summary.C = history.filter(r => r.status === 'C').length;
  summary.ST = history.filter(r => r.status === 'ST').length;
  summary.PG = history.filter(r => r.status === 'PG').length;
  summary.VC = history.filter(r => r.status === 'VC').length;

  // RESTORE PENTING: Kalkulasi Rekap Bulanan Asli Anda
  let sParts = startStr.split('-');
  let mmYYYY = sParts[1] + '/' + sParts[0];
  const sheetRekap = ss.getSheetByName('Rekap_Bulanan');
  if (sheetRekap) {
    const rekapData = sheetRekap.getDataRange().getDisplayValues();
    for(let i=1; i<rekapData.length; i++) {
      let row = rekapData[i];
      if(!row || !row[0] || !row[1]) continue;
      if(row[0].toString().trim() === mmYYYY && row[1].toString().trim() === nrpp.toString().trim()) {
        summary.S = Math.max(summary.S, parseInt(row[5]) || 0);
        summary.C = Math.max(summary.C, parseInt(row[6]) || 0);
        summary.PG = Math.max(summary.PG, parseInt(row[7]) || 0);
        summary.ST = Math.max(summary.ST, parseInt(row[8]) || 0);
        summary.VC = Math.max(summary.VC, parseInt(row[9]) || 0);
        break;
      }
    }
  }

  let totalHariKerja = 0;
  for (let d = new Date(startStr); d <= new Date(endStr); d.setDate(d.getDate() + 1)) {
     let raw = Utilities.formatDate(d, "GMT+7", "dd/MM/yyyy");
     if (d.getDay() !== 0 && d.getDay() !== 6 && !liburMap[raw]) totalHariKerja++;
  }
  summary.TotalHariKerja = totalHariKerja;

  return { profil: profil, data: history, summary: summary, liburData: liburMap };
}

function getVariabelGajiDataByDate(start, end, nrppFilter, ptFilter) {
  const ss = getDbInstance();
  const raw = ss.getSheetByName('Raw_Kehadiran').getDataRange().getDisplayValues();
  const master = ss.getSheetByName('Master_Karyawan').getDataRange().getDisplayValues();
  let mLookup = {};
  for(let i=1; i<master.length; i++) {
    if(master[i] && master[i][0]) mLookup[master[i][0].toString().trim()] = { nama: master[i][1], dept: master[i][5], pt: master[i][7] };
  }

  let result = {};
  for(let i=1; i<raw.length; i++) {
    let row = raw[i];
    if(!row || !row[0] || !row[1]) continue;
    let tgl = row[0].toString().trim(); 
    let nrpp = row[1].toString().trim();
    if(isDateInRange(tgl, start, end)) {
      let m = mLookup[nrpp];
      if(!m) continue;
      if(ptFilter !== 'ALL' && m.pt !== ptFilter) continue;
      if(nrppFilter !== 'ALL' && nrpp !== nrppFilter) continue;

      if(!result[nrpp]) result[nrpp] = { nrpp: nrpp, nama: m.nama, dept: m.dept, totalHadir: 0 };
      if((row[5]||"H").toString().trim().toUpperCase() === 'H') result[nrpp].totalHadir++;
    }
  }
  return Object.values(result).map(r => ({ nrpp: r.nrpp, nama: r.nama, dept: r.dept, totalHadir: r.totalHadir }));
}

/**
 * FUNGSI BARU: Mengubah HTML Pratinjau menjadi PDF Base64
 * Diletakkan di 07_LaporanService.gs
 * 
 * @param {string} htmlContent - Isi HTML dari area pratinjau
 * @param {string} filename - Nama file PDF
 * @returns {object} {success: boolean, base64: string, message: string}
 */
function konversiKePDFBase64(htmlContent, filename) {
  try {
    // Tambahkan sedikit CSS bawaan agar tabel pratinjau rapi saat jadi PDF
    const htmlUntukPDF = `
      <html>
        <head>
          <style>
            body { font-family: sans-serif; padding: 15px; }
            table { width: 100%; border-collapse: collapse; margin-top: 10px; font-size: 12px; }
            th, td { border: 1px solid #000; padding: 6px; text-align: left; }
            th { background-color: #f2f2f2; font-weight: bold; }
          </style>
        </head>
        <body>
          ${htmlContent}
        </body>
      </html>
    `;
    
    // Proses konversi oleh mesin Google (Aman dari V8 Error)
    const blob = Utilities.newBlob(htmlUntukPDF, MimeType.HTML).getAs(MimeType.PDF);
    const base64String = Utilities.base64Encode(blob.getBytes());
    
    return { 
      success: true, 
      base64: base64String, 
      message: "PDF berhasil disandikan."
    };
    
  } catch (error) {
    Logger.log("ERROR konversiKePDFBase64: " + error.message);
    return { 
      success: false, 
      message: error.message 
    };
  }
}

// =========================================================================
// DATA REKAP KEHADIRAN UNTUK VARIABEL GAJI
// =========================================================================
function getRekapKehadiranGaji(periode) {
  try {
    // 1. Panggil ID Database dari 01_Config.gs agar dinamis dan aman
    var idDb = appConfig().DB_UTAMA;
    var ss = SpreadsheetApp.openById(idDb);
    
    // 2. Ambil data dari sheet (Sesuaikan nama sheet dengan database Anda)
    var sheetMaster = ss.getSheetByName("Master_Karyawan");
    var sheetRekap = ss.getSheetByName("REKAP_BULANAN"); // Ganti jika namanya beda
    
    // --- AREA LOGIKA PENGAMBILAN DATA ---
    // Di sinilah Anda melooping data karyawan dari sheetMaster
    // dan mencocokkannya dengan total kehadiran di sheetRekap.
    
    // Contoh format data mutlak yang WAJIB dikembalikan ke depan (Frontend):
    var dataHasil = [
      // Format array object ini tidak boleh diubah agar DataTables tidak error
      { nrpp: "10017031", nama: "RANGGA GUMILANG", gol: "III-A", jabatan: "Sales Senior", departemen: "After Sales", totalKehadiran: 10 },
      { nrpp: "10000117", nama: "SUSENO", gol: "IV-A", jabatan: "Sales Senior", departemen: "After Sales", totalKehadiran: 14 }
    ];
    
    // Nanti ganti 'dataHasil' di atas dengan array hasil looping asli dari sheet Anda.
    // Pastikan logika totalKehadiran menggunakan rumus: H = H + ST + VC - PG - S - C
    
    return dataHasil; 
    
  } catch (error) {
    return { error: error.message };
  }
}
