function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('KeuanganKu - Multi User')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// FUNGSI BARU: Menerima 'username' untuk menentukan Sheet mana yang dibuka
function getSheet(username) {
  // Jika nama kosong, gunakan 'Data_Umum'
  const sheetName = (username && username.trim() !== "") ? username.trim() : "Data_Umum";
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  
  // Jika nama sheet belum ada, otomatis BUAT SHEET BARU
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(["ID", "Tanggal", "Jenis", "Keterangan", "Nominal"]);
    sheet.getRange("A1:E1").setFontWeight("bold").setBackground("#f8f9fa");
  }
  return sheet;
}

// Menerima parameter username dari web
function getTransactions(username) {
  const sheet = getSheet(username);
  const data = sheet.getDataRange().getValues();
  let result = [];
  
  for (let i = 1; i < data.length; i++) {
    result.push({
      id: data[i][0], date: data[i][1], type: data[i][2], text: data[i][3], amount: data[i][4]
    });
  }
  return result.reverse(); 
}

function addTransaction(data, username) {
  const sheet = getSheet(username);
  const id = new Date().getTime(); 
  const date = new Date().toLocaleDateString('id-ID');
  
  sheet.appendRow([id, date, data.type, data.text, Number(data.amount)]);
  return getTransactions(username);
}

function deleteTransaction(id, username) {
  const sheet = getSheet(username);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id) {
      sheet.deleteRow(i + 1); 
      break;
    }
  }
  return getTransactions(username); 
}
