function getLiburData() {
  try {
    const sheet = getDbInstance().getSheetByName('Master_Libur');
    if(!sheet) return [];
    const data = sheet.getDataRange().getDisplayValues();
    let hasil = [];
    for(let i=1; i<data.length; i++) {
      if(data[i][0]) hasil.push({tgl: data[i][0], status: data[i][1], ket: data[i][2]});
    }
    return hasil;
  } catch(e) { return []; }
}

function saveLiburData(tglInput, status, ket) {
  try {
    const sheet = getDbInstance().getSheetByName('Master_Libur');
    const data = sheet.getDataRange().getDisplayValues();
    let parts = tglInput.split('-'); let tglFormat = `${parts[2]}/${parts[1]}/${parts[0]}`;
    
    let found = -1;
    for(let i=1; i<data.length; i++) {
      if(data[i][0] === tglFormat) { found = i+1; break; }
    }
    if(found > -1) {
      sheet.getRange(found, 2).setValue(status);
      sheet.getRange(found, 3).setValue(ket);
    } else {
      sheet.appendRow([tglFormat, status, ket]);
    }
    return {success: true, message: "Hari Libur berhasil disimpan."};
  } catch(e) { return {success: false, message: e.message}; }
}

function deleteLiburData(tglInput) {
  try {
    const sheet = getDbInstance().getSheetByName('Master_Libur');
    const data = sheet.getDataRange().getDisplayValues();
    let parts = tglInput.split('-'); let tglFormat = `${parts[2]}/${parts[1]}/${parts[0]}`;
    
    for(let i=1; i<data.length; i++) {
      if(data[i][0] === tglFormat) { 
        sheet.deleteRow(i+1); 
        return {success: true, message: "Hari Libur dihapus."}; 
      }
    }
    return {success: false, message: "Data tidak ditemukan."};
  } catch(e) { return {success: false, message: e.message}; }
}
