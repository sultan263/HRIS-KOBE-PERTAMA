function getWorkingHours(dateStr) {
  if(!dateStr) return appConfig().NORMAL_SHIFT.MON_THU;
  const config = appConfig();
  const parts = dateStr.toString().split('/');
  if(parts.length !== 3) return config.NORMAL_SHIFT.MON_THU;
  
  const dateObj = new Date(parts[2], parts[1] - 1, parts[0]);
  const dayOfWeek = dateObj.getDay(); 

  const ramadanStart = new Date(config.RAMADAN_PERIOD.START);
  const ramadanEnd = new Date(config.RAMADAN_PERIOD.END);
  const isRamadan = (dateObj >= ramadanStart && dateObj <= ramadanEnd);

  const shiftData = isRamadan ? config.RAMADAN_SHIFT : config.NORMAL_SHIFT;
  return (dayOfWeek === 5) ? shiftData.FRI : shiftData.MON_THU;
}

function convertSmartDate(value) {
  if (!value) return value;
  if (typeof value === 'string' && (value.includes('/') || value.includes('-'))) return value;
  let serial = parseFloat(value);
  if (!isNaN(serial)) {
    const excelDate = new Date((serial - 25569) * 86400 * 1000);
    return Utilities.formatDate(excelDate, "GMT+7", "dd/MM/yyyy");
  }
  return value;
}

// FUNGSI BARU: Sabuk Pengaman Konversi Jam Mentah (Pecahan) Excel
function convertSmartTime(value) {
  if (value === undefined || value === null || value === "") return "";
  
  // Jika string dan mengandung titik dua, format standar
  if (typeof value === 'string' && value.includes(':')) {
    let match = value.match(/(\d{1,2}):(\d{1,2}):?(\d{1,2})?/);
    if (match) {
      let h = match[1].padStart(2, '0'); let m = match[2].padStart(2, '0'); let s = (match[3] || "00").padStart(2, '0');
      return `${h}:${m}:${s}`;
    }
    return value.trim();
  }

  // Jika format berupa serial pecahan desimal Excel (contoh: 0.327777 -> 07:52:00)
  let serial = parseFloat(value);
  if (!isNaN(serial)) {
    let fraction = serial;
    // Jika terdapat bagian tanggal (nilai >= 1), ambil pecahannya saja
    if (serial >= 1) fraction = serial - Math.floor(serial);
    
    let totalSeconds = Math.round(fraction * 86400); // 86400 detik dalam 1 hari
    let h = Math.floor(totalSeconds / 3600);
    let m = Math.floor((totalSeconds % 3600) / 60);
    let s = totalSeconds % 60;
    return `${h.toString().padStart(2, '0')}:${m.toString().padStart(2, '0')}:${s.toString().padStart(2, '0')}`;
  }
  return value.toString().trim();
}

function timeToSeconds(timeStr) {
  if(!timeStr || typeof timeStr !== 'string') return 0;
  let parts = timeStr.trim().split(':');
  if(parts.length < 2) return 0;
  return (parseInt(parts[0]) || 0) * 3600 + (parseInt(parts[1]) || 0) * 60 + (parseInt(parts[2]) || 0);
}

function secondsToTimeStr(secs) {
  if (secs <= 0) return "00:00:00";
  let h = Math.floor(secs / 3600); let m = Math.floor((secs % 3600) / 60); let s = secs % 60;
  return `${h.toString().padStart(2,'0')}:${m.toString().padStart(2,'0')}:${s.toString().padStart(2,'0')}`;
}

function formatTimeForInput(timeStr) {
  if (!timeStr || timeStr === "-" || timeStr === "") return "";
  let match = timeStr.toString().match(/(\d{1,2}):(\d{1,2}):?(\d{1,2})?/);
  if (!match) return "";
  let h = match[1].padStart(2, '0'); let m = match[2].padStart(2, '0'); let s = (match[3] || "00").padStart(2, '0');
  return `${h}:${m}:${s}`;
}

function getWorkdaysByDate(startStr, endStr) {
  if(!startStr || !endStr) return 0;
  let s = new Date(startStr); let e = new Date(endStr);
  let count = 0;
  for (let d = new Date(s); d <= e; d.setDate(d.getDate() + 1)) {
    if(d.getDay() !== 0 && d.getDay() !== 6) count++;
  }
  return count;
}

function isDateInRange(tglSheet, startStr, endStr) {
  if(!tglSheet || !startStr || !endStr) return false;
  let p = tglSheet.toString().split('/');
  if(p.length !== 3) return false;
  let yyyyMmDd = `${p[2]}-${p[1].padStart(2,'0')}-${p[0].padStart(2,'0')}`;
  return (yyyyMmDd >= startStr && yyyyMmDd <= endStr);
}
