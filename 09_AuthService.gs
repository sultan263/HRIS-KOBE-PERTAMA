/**
 * Fungsi untuk memverifikasi login user
 */
function checkLogin(username, password) {
  try {
    const ss = getDbInstance();
    const sheet = ss.getSheetByName('Master_User');
    if (!sheet) return { success: false, message: "Database User tidak ditemukan!" };
    
    const data = sheet.getDataRange().getDisplayValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === username && data[i][1] === password) {
        return {
          success: true,
          userData: {
            username: data[i][0],
            nama: data[i][2],
            role: data[i][3]
          }
        };
      }
    }
    return { success: false, message: "Username atau Password salah!" };
  } catch (e) {
    return { success: false, message: e.message };
  }
}
