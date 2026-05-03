function processExcelUpload(dataRows) {
  try {
    const ss = getDbInstance();
    const sheetRaw = ss.getSheetByName('Raw_Kehadiran');
    const sheetMaster = ss.getSheetByName('Master_Karyawan');
    const masterData = sheetMaster.getDataRange().getValues();
    const daftarNrppMaster = masterData.map(row => row[0].toString().trim());

    let dataSiapTulis = []; let jumlahDiskip = 0;
    for (let i = 1; i < dataRows.length; i++) {
      let row = dataRows[i];
      if (!row || row.length < 2 || !row[0]) continue; 
      let nrppExcel = row[0].toString().trim();
      if (!daftarNrppMaster.includes(nrppExcel)) { jumlahDiskip++; continue; }
      
      let nama = row[1] ? row[1].toString().trim() : "";
      
      // PERBAIKAN: Melewatkan raw jam masuk & keluar ke mesin konversi SmartTime
      let signIn = convertSmartTime(row[2]);
      let signOut = convertSmartTime(row[3]);
      
      let status = row[4] ? row[4].toString().trim() : "H";
      let rawDate = row[5] ? row[5] : "";
      let tanggalRapih = convertSmartDate(rawDate);
      dataSiapTulis.push([tanggalRapih, nrppExcel, nama, signIn, signOut, status]);
    }

    if (dataSiapTulis.length > 0) {
      sheetRaw.getRange(sheetRaw.getLastRow() + 1, 1, dataSiapTulis.length, 6).setValues(dataSiapTulis);
      updateRekapanBulanan(dataSiapTulis); 
      let msg = `${dataSiapTulis.length} data berhasil masuk.`;
      if(jumlahDiskip > 0) msg += ` (${jumlahDiskip} karyawan otomatis difilter).`;
      return { success: true, message: msg };
    } else { return { success: false, message: "Tidak ada data yang cocok dengan Master Karyawan." }; }
  } catch (error) { return { success: false, message: error.message }; }
}

function updateRekapanBulanan(dataBaru) {
  const sheetRekap = getDbInstance().getSheetByName('Rekap_Bulanan');
  const sheetMaster = getDbInstance().getSheetByName('Master_Karyawan');
  let masterValues = sheetMaster.getDataRange().getDisplayValues();
  let kamusDivisi = {};
  for(let i=1; i<masterValues.length; i++) { 
    if(masterValues[i] && masterValues[i][0]) { kamusDivisi[masterValues[i][0].toString().trim()] = masterValues[i][6] || "-"; }
  }
  
  let hitungHadir = {};
  dataBaru.forEach(row => {
    let tgl = row[0]; if(!tgl || typeof tgl !== 'string') return;
    let bulanTahun = tgl.substring(3, 10); let nrpp = row[1]; let nama = row[2];
    let key = bulanTahun + "_" + nrpp;
    if(!hitungHadir[key]) { hitungHadir[key] = { bulan: bulanTahun, nrpp: nrpp, nama: nama, divisi: (kamusDivisi[nrpp]||"-"), hadir: 0 }; }
    if(row[5] === 'H') hitungHadir[key].hadir += 1;
  });
  
  let barisBaru = [];
  for (let key in hitungHadir) {
    let d = hitungHadir[key];
    barisBaru.push([d.bulan, d.nrpp, d.nama, d.divisi, d.hadir, 0, 0, 0, 0, 0]);
  }
  if(barisBaru.length > 0) sheetRekap.getRange(sheetRekap.getLastRow() + 1, 1, barisBaru.length, 10).setValues(barisBaru);
}

function getAttendanceForEdit(dateStr, ptFilter) {
  const ss = getDbInstance();
  const masterData = ss.getSheetByName('Master_Karyawan').getDataRange().getDisplayValues();
  let employees = [];
  for(let i=1; i<masterData.length; i++){
    let row = masterData[i];
    if(!row[0]) continue;
    if(ptFilter === "ALL" || row[7] === ptFilter) {
      employees.push({nrpp: row[0].trim(), nama: row[1]});
    }
  }
  
  let parts = dateStr.split('-');
  let tglFormatted = `${parts[2]}/${parts[1]}/${parts[0]}`;
  const rawData = ss.getSheetByName('Raw_Kehadiran').getDataRange().getDisplayValues();
  
  let attMap = {};
  for(let i=1; i<rawData.length; i++){
    if(rawData[i][0] === tglFormatted) {
       attMap[rawData[i][1].trim()] = { 
         masuk: formatTimeForInput(rawData[i][3]), 
         keluar: formatTimeForInput(rawData[i][4]), 
         status: rawData[i][5] 
       };
    }
  }
  
  employees.forEach(emp => {
    let att = attMap[emp.nrpp];
    emp.masuk = att ? att.masuk : "";
    emp.keluar = att ? att.keluar : "";
    emp.status = att ? att.status : "-";
  });
  return employees;
}

function saveAbsensiManual(tglInput, nrpp, masuk, keluar, status) {
  try {
    const ss = getDbInstance();
    const sheetRaw = ss.getSheetByName('Raw_Kehadiran');
    const rawData = sheetRaw.getDataRange().getDisplayValues();
    let parts = tglInput.split('-'); let tglFormatted = `${parts[2]}/${parts[1]}/${parts[0]}`;
    const masterData = ss.getSheetByName('Master_Karyawan').getDataRange().getDisplayValues();
    let nama = "-";
    for(let i=1; i<masterData.length; i++) {
      let row = masterData[i];
      if(row && row[0] && row[0].toString().trim() === nrpp.toString().trim()) { nama = row[1] || "-"; break; }
    }
    let foundRowIndex = -1;
    for(let i=1; i<rawData.length; i++) {
      let row = rawData[i];
      if(row && row[0] === tglFormatted && row[1].toString().trim() === nrpp.toString().trim()) { foundRowIndex = i + 1; break; }
    }
    if(foundRowIndex > -1) {
      sheetRaw.getRange(foundRowIndex, 4).setValue(masuk);
      sheetRaw.getRange(foundRowIndex, 5).setValue(keluar);
      sheetRaw.getRange(foundRowIndex, 6).setValue(status);
    } else { sheetRaw.appendRow([tglFormatted, nrpp, nama, masuk, keluar, status]); }
    return { success: true, message: `Disimpan` };
  } catch(e) { return { success: false, message: e.message }; }
}

function deleteAbsensiManual(tglInput, nrpp) {
  try {
    const sheetRaw = getDbInstance().getSheetByName('Raw_Kehadiran');
    const rawData = sheetRaw.getDataRange().getDisplayValues();
    let parts = tglInput.split('-'); let tglFormatted = `${parts[2]}/${parts[1]}/${parts[0]}`;
    let foundRowIndex = -1;
    for(let i=1; i<rawData.length; i++) {
      if(rawData[i][0] === tglFormatted && rawData[i][1].toString().trim() === nrpp.toString().trim()) { foundRowIndex = i + 1; break; }
    }
    if(foundRowIndex > -1) {
      sheetRaw.deleteRow(foundRowIndex);
      return { success: true, message: "Dihapus" };
    }
    return { success: false, message: "Tidak ada data" };
  } catch(e) { return { success: false, message: e.message }; }
}
