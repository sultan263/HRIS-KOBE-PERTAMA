function getDaftarKaryawan() {
  try {
    const sheet = getDbInstance().getSheetByName('Master_Karyawan');
    if(!sheet) return { error: "Sheet Master_Karyawan tidak ditemukan." };
    const data = sheet.getDataRange().getDisplayValues();
    let hasil = []; let ptList = [];
    for(let i=1; i<data.length; i++) {
      let row = data[i];
      if(row && row[0]) {
        let pt = (row[7] || "").toString().trim();
        hasil.push({nrpp: row[0].toString().trim(), nama: row[1] || "-", pt: pt});
        if(pt && !ptList.includes(pt)) ptList.push(pt);
      }
    }
    if(ptList.length === 0) ptList.push("PT. KOBXINDO EQUIPMENT");
    return { success: true, data: hasil, perusahaanList: ptList };
  } catch (err) { return { error: err.message }; }
}
