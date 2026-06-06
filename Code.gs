/**
 * Saku Ku - Google Sheets Backend Engine
 * Berfungsi sebagai REST API (CRUD) untuk aplikasi Manajemen Anggaran Saku Ku.
 */

// Global helper untuk mendapatkan spreadsheet aktif secara aman
function getActiveSpreadsheetSecurely() {
  var ss = null;
  try {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  } catch (e) {
    // Terjadi jika dijalankan di luar konteks dokumen aktif
  }
  
  if (!ss) {
    var userProperties = PropertiesService.getUserProperties();
    var savedId = userProperties.getProperty('SAKUKU_SPREADSHEET_ID');
    
    if (savedId) {
      try {
        ss = SpreadsheetApp.openById(savedId);
      } catch (e) {
        userProperties.deleteProperty('SAKUKU_SPREADSHEET_ID');
      }
    }
    
    if (!ss) {
      ss = SpreadsheetApp.create("Saku Ku - Database Anggaran Bulanan");
      userProperties.setProperty('SAKUKU_SPREADSHEET_ID', ss.getId());
    }
  }
  return ss;
}

// Menangani permintaan GET (Bisa render halaman Index.html atau merespons API JSON)
function doGet(e) {
  // Jika parameter api=true dipanggil, kembalikan respons JSON data anggaran (READ)
  if (e && e.parameter && e.parameter.api === "true") {
    try {
      setupSpreadsheet();
      var data = readDataFromSheet();
      return handleResponse({
        status: "success",
        message: "Data berhasil diunduh dari Spreadsheet.",
        data: data
      });
    } catch (err) {
      return handleResponse({
        status: "error",
        message: "Gagal membaca data: " + err.toString()
      });
    }
  }

  // Fallback: Sajikan file interface Index.html secara langsung dari Web App Apps Script
  setupSpreadsheet();
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Saku Ku - Asisten Anggaran Bulanan')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// Menangani permintaan POST (Menulis / Menyimpan Data / UPDATE / CREATE)
function doPost(e) {
  try {
    setupSpreadsheet();
    var requestData = JSON.parse(e.postData.contents);
    var action = requestData.action;
    
    if (action === "sync") {
      writeDataToSheet(requestData.data);
      return handleResponse({
        status: "success",
        message: "Data anggaran berhasil disinkronisasikan ke Spreadsheet!"
      });
    }
    
    return handleResponse({
      status: "error",
      message: "Aksi '" + action + "' tidak dikenali oleh sistem."
    });
  } catch (err) {
    return handleResponse({
      status: "error",
      message: "Gagal memproses permintaan POST: " + err.toString()
    });
  }
}

// Helper untuk membungkus respons JSON dengan header CORS yang tepat
function handleResponse(responseObject) {
  return ContentService.createTextOutput(JSON.stringify(responseObject))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * 1. FUNGSI SETUP SPREADSHEET (Auto-Create Sheets & Headers)
 * Otomatis mendeteksi dan membuat struktur database sheet jika belum ada.
 */
function setupSpreadsheet() {
  var ss = getActiveSpreadsheetSecurely();
  
  // A. Inisialisasi Sheet "Konfigurasi"
  var configSheet = ss.getSheetByName("Konfigurasi");
  if (!configSheet) {
    configSheet = ss.insertSheet("Konfigurasi");
    configSheet.appendRow(["Gaji / Pendapatan Bulanan"]);
    configSheet.appendRow([7500000]); // Nilai awal default
    
    // Formatting header
    configSheet.getRange("A1").setFontWeight("bold").setBackground("#EEF2FF").setFontColor("#4F46E5");
    configSheet.getRange("A2").setNumberFormat("Rp#,##0");
    configSheet.autoResizeColumn(1);
  }
  
  // B. Inisialisasi Sheet "Kategori"
  var catSheet = ss.getSheetByName("Kategori");
  if (!catSheet) {
    catSheet = ss.insertSheet("Kategori");
    catSheet.appendRow(["ID Kategori", "Nama Kategori"]);
    
    // Baris data awal dummy
    catSheet.appendRow(["g1", "Kebutuhan Pokok"]);
    catSheet.appendRow(["g2", "Gaya Hidup & Hiburan"]);
    
    // Formatting header
    catSheet.getRange("A1:B1").setFontWeight("bold").setBackground("#EEF2FF").setFontColor("#4F46E5");
    catSheet.autoResizeColumns(1, 2);
  }
  
  // C. Inisialisasi Sheet "Item_Anggaran"
  var itemSheet = ss.getSheetByName("Item_Anggaran");
  if (!itemSheet) {
    itemSheet = ss.insertSheet("Item_Anggaran");
    itemSheet.appendRow(["ID Item", "ID Kategori", "Nama Item", "Harian (TRUE/FALSE)", "Tarif Harian", "Override Hari", "Tunai", "Non Tunai"]);
    
    // Baris data awal dummy
    itemSheet.appendRow(["i1", "g1", "Belanja Dapur Mingguan", false, 0, "", 1500000, 0]);
    itemSheet.appendRow(["i2", "g1", "Uang Makan Siang Kerja", true, 35000, 20, 0, 700000]);
    itemSheet.appendRow(["i3", "g2", "Langganan Netflix & Spotify", false, 0, "", 0, 230000]);
    
    // Formatting header
    itemSheet.getRange("A1:H1").setFontWeight("bold").setBackground("#EEF2FF").setFontColor("#4F46E5");
    itemSheet.getRange("E2:E").setNumberFormat("Rp#,##0");
    itemSheet.getRange("G2:H").setNumberFormat("Rp#,##0");
    itemSheet.autoResizeColumns(1, 8);
  }
  
  var defaultSheet = ss.getSheetByName("Sheet1");
  if (defaultSheet && defaultSheet.getLastRow() === 0) {
    ss.deleteSheet(defaultSheet);
  }
}

/**
 * 2. FUNGSI READ (Mengambil data dari Spreadsheet dan dikonversi ke JSON Saku Ku)
 */
function readDataFromSheet() {
  var ss = getActiveSpreadsheetSecurely();
  
  // Membaca Nominal Gaji
  var configSheet = ss.getSheetByName("Konfigurasi");
  var gaji = Number(configSheet.getRange(2, 1).getValue()) || 0;
  
  // Membaca Kategori (Groups)
  var catSheet = ss.getSheetByName("Kategori");
  var catRows = catSheet.getDataRange().getValues();
  var groups = [];
  
  for (var i = 1; i < catRows.length; i++) {
    var groupId = catRows[i][0].toString().trim();
    var groupName = catRows[i][1].toString().trim();
    
    if (groupId) {
      groups.push({
        id: groupId,
        name: groupName,
        items: []
      });
    }
  }
  
  // Membaca Item Anggaran
  var itemSheet = ss.getSheetByName("Item_Anggaran");
  var itemRows = itemSheet.getDataRange().getValues();
  
  for (var j = 1; j < itemRows.length; j++) {
    var itemId = itemRows[j][0].toString().trim();
    var groupId = itemRows[j][1].toString().trim();
    var itemName = itemRows[j][2].toString().trim();
    var isDaily = itemRows[j][3] === true || itemRows[j][3].toString().toUpperCase() === "TRUE";
    var dailyRate = Number(itemRows[j][4]) || 0;
    
    var daysOverride = itemRows[j][5] === "" ? 0 : Number(itemRows[j][5]);
    
    var tunai = Number(itemRows[j][6]) || 0;
    var non_tunai = Number(itemRows[j][7]) || 0;
    
    if (itemId && groupId) {
      var parentGroup = groups.find(function(g) {
        return g.id === groupId;
      });
      
      if (parentGroup) {
        parentGroup.items.push({
          id: itemId,
          name: itemName,
          isDaily: isDaily,
          dailyRate: dailyRate,
          daysOverride: daysOverride,
          tunai: tunai,
          non_tunai: non_tunai
        });
      }
    }
  }
  
  return {
    gaji: gaji,
    groups: groups
  };
}

/**
 * 3. FUNGSI WRITE/WRITE BACK (Menulis ulang database Spreadsheet berdasarkan state JSON Saku Ku)
 */
function writeDataToSheet(data) {
  if (!data || typeof data.gaji === 'undefined') {
    throw new Error("Struktur data tidak valid untuk disimpan.");
  }
  
  var ss = getActiveSpreadsheetSecurely();
  
  var configSheet = ss.getSheetByName("Konfigurasi");
  configSheet.clear();
  configSheet.appendRow(["Gaji / Pendapatan Bulanan"]);
  configSheet.appendRow([Number(data.gaji) || 0]);
  
  var catSheet = ss.getSheetByName("Kategori");
  catSheet.clear();
  catSheet.appendRow(["ID Kategori", "Nama Kategori"]);
  
  var itemSheet = ss.getSheetByName("Item_Anggaran");
  itemSheet.clear();
  itemSheet.appendRow(["ID Item", "ID Kategori", "Nama Item", "Harian (TRUE/FALSE)", "Tarif Harian", "Override Hari", "Tunai", "Non Tunai"]);
  
  if (data.groups && data.groups.length > 0) {
    data.groups.forEach(function(group) {
      if (group.id && group.name) {
        catSheet.appendRow([group.id, group.name]);
        
        if (group.items && group.items.length > 0) {
          group.items.forEach(function(item) {
            itemSheet.appendRow([
              item.id,
              group.id,
              item.name,
              item.isDaily,
              item.dailyRate,
              item.daysOverride === 0 ? "" : item.daysOverride,
              item.tunai,
              item.non_tunai
            ]);
          });
        }
      }
    });
  }
  
  // Format ulang style visual
  configSheet.getRange("A1").setFontWeight("bold").setBackground("#EEF2FF").setFontColor("#4F46E5");
  configSheet.getRange("A2").setNumberFormat("Rp#,##0");
  
  catSheet.getRange("A1:B1").setFontWeight("bold").setBackground("#EEF2FF").setFontColor("#4F46E5");
  itemSheet.getRange("A1:H1").setFontWeight("bold").setBackground("#EEF2FF").setFontColor("#4F46E5");
  itemSheet.getRange("E2:E").setNumberFormat("Rp#,##0");
  itemSheet.getRange("G2:H").setNumberFormat("Rp#,##0");
  
  configSheet.autoResizeColumn(1);
  catSheet.autoResizeColumns(1, 2);
  itemSheet.autoResizeColumns(1, 8);
}
