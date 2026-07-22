// 1. ROUTING HALAMAN UTAMA WEB APP
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('MANRISK - Loka POM Kobar')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 2. FUNGSI UTAMA: AMBIL DATA DARI TAB DAFTARRESIKO
function getDaftarRisiko() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("DaftarResiko");
    
    // Jika sheet belum ada, buat otomatis agar tidak memicu eror crash
    if (!sheet) {
      return [];
    }
    
    var data = sheet.getDataRange().getValues();
    return data; 
  } catch (err) {
    Logger.log("Eror pada getDaftarRisiko: " + err.toString());
    return [];
  }
}

// 3. AMBIL DATA DARI TAB KONTEKS UNTUK DROPDOWN ENTRY RISIKO
// FUNGSI UNTUK MENGAMBIL DAFTAR SASARAN/KONTEKS UNTUK DROPDOWN ENTRY RISIKO
// FUNGSI UNTUK MENGAMBIL DAFTAR KONTEKS SECARA DINAMIS BERDASARKAN HEADER
// FUNGSI FIXED UNTUK MENYUPLAI DATA DROPDOWN SASARAN DI MENU ENTRY RISIKO
function getKonteksList() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Konteks");
    
    if (!sheet) {
      Logger.log("Error: Sheet 'Konteks' tidak ditemukan.");
      return [];
    }
    
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return []; // Hanya ada header, langsung return kosong
    
    var data = sheet.getDataRange().getValues();
    var headerRow = data[0]; // Baris pertama sebagai acuan nama kolom
    
    // Cari letak kolom yang berisi teks "konteks" secara otomatis (tidak sensitif huruf besar/kecil)
    var idxKonteks = headerRow.findIndex(function(cell) {
      return String(cell).toLowerCase().indexOf("konteks") !== -1;
    });
    
    // Jika tidak ditemukan lewat pencarian teks header, gunakan kolom ke-2 / B (indeks 1) sebagai standar
    if (idxKonteks === -1) {
      idxKonteks = 1;
    }
    
    var list = [];
    // Baca baris ke-2 sampai akhir
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (row && row.length > idxKonteks && row[idxKonteks]) {
        var nilaiKonteks = String(row[idxKonteks]).trim();
        if (nilaiKonteks !== "" && nilaiKonteks !== "-") {
          list.push({ konteks: nilaiKonteks });
        }
      }
    }
    return list;
  } catch (e) {
    Logger.log("Error pada getKonteksList: " + e.toString());
    return [];
  }
}

// 4. GENERATE KODE RISIKO OTOMATIS BERDASARKAN KATEGORI & KODE UNIT
function generateKodeRisiko(kategori, unitKode) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DaftarResiko");
    var data = sheet ? sheet.getDataRange().getValues() : [];
    
    var katMap = {
      "Risiko Strategis": "S", "Risiko Reputasi": "R", "Risiko Hukum": "H",
      "Risiko Sumber daya manusia": "SDM", "Risiko Operasional": "O", "Risiko IT": "IT",
      "Risiko kesehatan dan Keselamatan Kerja": "KK", "Risiko Fraud": "F",
      "Risiko Keuangan": "K", "Risiko Aset": "A"
    };
    var initialKat = katMap[kategori] || "X";
    
    var count = 0;
    // Cari kemunculan kode unit yang sama untuk menentukan nomor urut selanjutnya
    for (var i = 1; i < data.length; i++) {
      var existingKode = data[i][3] ? String(data[i][3]) : "";
      if (existingKode.includes("." + unitKode + ".")) {
        count++;
      }
    }
    
    var noUrut = String(count + 1).padStart(3, '0');
    return initialKat + "." + unitKode + "." + noUrut;
  } catch(e) {
    return "ERR.00.000";
  }
}

// 5. SIMPAN DATA IDENTIFIKASI RISIKO BARU (ENTRY RISIKO)
function simpanIdentifikasi(formData, token) {
  try {
    pastikanAkses_(token, ["admin", "operator"]);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("DaftarResiko");
    if (!sheet) return { success: false, message: "Sheet DaftarResiko tidak ditemukan!" };
    
    var timestamp = new Date();
    var lastRow = sheet.getLastRow();
    var displayNo = lastRow; 

    // Susun array 23 Kolom Kosong/Default terlebih dahulu agar struktur tabel tidak rusak
    var rowData = new Array(23).fill("");
    
    rowData[0] = displayNo;              // No.
    rowData[1] = formData.sasaran;       // Sasaran
    rowData[2] = formData.kegiatan;      // Kegiatan/ Proses Bisnis
    rowData[3] = formData.kodeRisiko;    // Kode Risiko
    rowData[4] = formData.kategori;      // Kategori Risiko
    rowData[5] = formData.riskEvent;     // Issue ISO - Risk Event / Uraian Peristiwa
    rowData[6] = formData.penyebab;      // Issue ISO - Penyebab Risiko
    rowData[7] = formData.sumber;        // Sumber Risiko
    rowData[8] = formData.akibat;        // Akibat/Potensi Kerugian
    rowData[9] = formData.pemilik;       // Pemilik Risiko
    rowData[10] = formData.unitKerja;    // Nama Unit Kerja Terkait
    rowData[21] = timestamp;             // Tanggal entri (Kolom ke-22 / V)

    sheet.appendRow(rowData);
    return { success: true, message: "Data berhasil disimpan dengan Kode: " + formData.kodeRisiko };
  } catch(err) {
    return { success: false, message: err.toString() };
  }
}

// 6. HITUNG LEVEL RISIKO DARI MATRIKS 5x5
function hitungLevelRisiko(kemungkinan, dampak) {
  var k = parseInt(kemungkinan);
  var d = parseInt(dampak);
  var matriks = {
    5: [9, 15, 18, 23, 25],
    4: [6, 12, 16, 19, 24],
    3: [4, 10, 14, 17, 22],
    2: [2, 7, 11, 13, 21],
    1: [1, 3, 5, 8, 20]
  };
  try {
    var skor = matriks[k][d - 1];
    var tingkat = "Rendah"; var warna = "#00a707";
    if (skor >= 6 && skor <= 12) { tingkat = "Sedang"; warna = "#fff714"; }
    else if (skor >= 13 && skor <= 19) { tingkat = "Tinggi"; warna = "#ff9f00"; }
    else if (skor >= 20) { tingkat = "Sangat Tinggi"; warna = "#ff001f"; }
    return { skor: skor, tingkat: tingkat, warna: warna };
  } catch(e) {
    return { skor: 0, tingkat: "Tidak Valid", warna: "#ffffff" };
  }
}


// 7. MENCARI RISIKO YANG BELUM DINILAI INHEREN-NYA (VERSI AMAN & FIXED)
function getRisikoBelumDinilai() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DaftarResiko");
    if (!sheet) return [];
    
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return []; // Jika hanya ada header saja, langsung return kosong
    
    var data = sheet.getDataRange().getValues();
    var list = [];
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      
      // Ambil nilai kolom L (indeks 11) untuk Kemungkinan Inheren
      // Gunakan pemeriksaan aman jika kolomnya belum terbentuk/undefined
      var kemungkinanInh = (row.length > 11) ? row[11] : "";
      
      // Jika kolom Kemungkinan Inheren (indeks 11) masih kosong atau strip (-)
      if (!kemungkinanInh || kemungkinanInh === "" || kemungkinanInh === "-") {
        list.push({
          rowNum: i + 1,
          kode: (row.length > 3) ? (row[3] || "N/A") : "N/A",
          kegiatan: (row.length > 2) ? (row[2] || "-") : "-",
          peristiwa: (row.length > 5) ? (row[5] || "-") : "-"
        });
      }
    }
    return list;
  } catch(e) {
    Logger.log("Error pada getRisikoBelumDinilai: " + e.toString());
    return [];
  }
}

// 8. SIMPAN PENILAIAN RISIKO INHEREN KEMBALI KE SHEET
function simpanAnalisisInheren(rowNum, kemungkinan, dampak, token) {
  try {
    pastikanAkses_(token, ["admin", "operator"]);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DaftarResiko");
    var hasil = hitungLevelRisiko(kemungkinan, dampak);
    
    sheet.getRange(rowNum, 12).setValue(kemungkinan); // Kolom L (Kemungkinan)
    sheet.getRange(rowNum, 13).setValue(dampak);      // Kolom M (Dampak)
    sheet.getRange(rowNum, 14).setValue(hasil.skor + " (" + hasil.tingkat + ")"); // Kolom N (Level)
    
    return { success: true, message: "Analisis Risiko Inheren Berhasil Disimpan!" };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

// ===================================================================
// 8B. AUTENTIKASI & MANAJEMEN SESI
// ===================================================================
// Struktur kolom sheet "Users": A=username, B=passwordHash, C=role, D=jabatan, E=salt

var SESSION_DURASI_DETIK = 21600; // 6 jam — batas maksimum CacheService

// Hasilkan salt acak unik per user
function buatSalt_() {
  return Utilities.getUuid();
}

// Hash password dengan SHA-256 + salt agar tidak tersimpan sebagai teks polos
function hashPassword_(password, salt) {
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(password) + String(salt));
  return digest.map(function(b) {
    var v = (b < 0) ? b + 256 : b;
    return ('0' + v.toString(16)).slice(-2);
  }).join('');
}

// LOGIN: verifikasi kredensial & buat token sesi baru
function loginUser(username, password) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
    if (!sheet) return { success: false, message: "Sheet Users tidak ditemukan!" };

    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (String(row[0]).trim() !== String(username).trim()) continue;

      var storedHash = row[1];
      var salt = row[4];

      if (!salt) {
        // Kompatibilitas mundur: user lama yang password-nya masih tersimpan polos.
        // Jika cocok, migrasikan otomatis ke hash+salt saat itu juga.
        if (String(storedHash) !== String(password)) {
          return { success: false, message: "Username atau password salah." };
        }
        var saltBaru = buatSalt_();
        var hashBaru = hashPassword_(password, saltBaru);
        sheet.getRange(i + 1, 2).setValue(hashBaru);
        sheet.getRange(i + 1, 5).setValue(saltBaru);
      } else {
        if (hashPassword_(password, salt) !== storedHash) {
          return { success: false, message: "Username atau password salah." };
        }
      }

      var role = row[2];
      var jabatan = row[3];
      var token = Utilities.getUuid();

      CacheService.getScriptCache().put(token, JSON.stringify({
        username: username, role: role, jabatan: jabatan
      }), SESSION_DURASI_DETIK);

      return { success: true, token: token, username: username, role: role, jabatan: jabatan };
    }
    return { success: false, message: "Username atau password salah." };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// Ambil ulang data sesi dari token (dipakai client saat refresh halaman)
function getSessionInfo(token) {
  try {
    var cached = CacheService.getScriptCache().get(token);
    if (!cached) return { success: false, message: "Sesi berakhir, silakan login kembali." };
    return { success: true, data: JSON.parse(cached) };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// LOGOUT: hapus token dari cache
function logoutUser(token) {
  try {
    if (token) CacheService.getScriptCache().remove(token);
    return { success: true };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// Helper otorisasi: lempar error jika token tidak valid / role tidak diizinkan.
// allowedRoles: array string, dicocokkan sebagai substring (role user bisa gabungan, mis. "Operator,admin")
function pastikanAkses_(token, allowedRoles) {
  var cached = token ? CacheService.getScriptCache().get(token) : null;
  if (!cached) throw new Error("Sesi tidak valid atau sudah berakhir. Silakan login kembali.");
  var session = JSON.parse(cached);
  var rolesUser = String(session.role || "").toLowerCase();
  var diizinkan = allowedRoles.some(function(r) { return rolesUser.indexOf(r.toLowerCase()) !== -1; });
  if (!diizinkan) throw new Error("Anda tidak memiliki akses untuk melakukan aksi ini.");
  return session;
}

// 9. MANAGEMENT USER: AMBIL SEMUA USER (DEFAULT DI ADMIN PANEL)
function getAllUsers(token) {
  try {
    pastikanAkses_(token, ["admin"]);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    var usersList = [];
    for (var i = 1; i < data.length; i++) {
      usersList.push({
        username: data[i][0],
        role: data[i][2],
        jabatan: data[i][3]
      });
    }
    return usersList;
  } catch(e) {
    Logger.log("Eror pada getAllUsers: " + e.toString());
    return [];
  }
}

// 10. SIMPAN USER BARU KE DATABASE USERS (ADMIN ONLY)
function simpanUser(userData, token) {
  try {
    pastikanAkses_(token, ["admin"]);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
    if (!sheet) return { success: false, message: "Sheet Users tidak ditemukan!" };

    var salt = buatSalt_();
    var hash = hashPassword_(userData.password, salt);
    sheet.appendRow([userData.username, hash, userData.roles.join(","), userData.jabatan, salt]);
    return { success: true, message: "User baru berhasil didaftarkan!" };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

// 11. HAPUS RISIKO (ADMIN ONLY)
function hapusRisiko(rowNum, token) {
  try {
    pastikanAkses_(token, ["admin"]);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DaftarResiko");
    if (!sheet) return { success: false, message: "Sheet tidak ditemukan!" };
    sheet.deleteRow(rowNum);
    return { success: true, message: "Data risiko berhasil dihapus!" };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

// 12. Mengambil semua data risiko untuk kebutuhan menu Analisis Risiko (Belum dinilai & Sudah dinilai)
// FUNGSI UTAMA: MENGAMBIL DAN MENGELOMPOKKAN DATA UNTUK MENU ANALISIS RISIKO
function getDataUntukAnalisis() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("DaftarResiko");
    
    if (!sheet) {
      Logger.log("Error: Sheet 'DaftarResiko' tidak ditemukan.");
      return { belumDinilai: [], sudahDinilai: [] };
    }
    
    var lastRow = sheet.getLastRow();
    if (lastRow <= 2) { // Menyesuaikan jika header Anda memakan 2 baris
      return { belumDinilai: [], sudahDinilai: [] };
    }
    
    var data = sheet.getDataRange().getValues();
    
    // BARIS DETERMINASI INDEKS (Mencari posisi kolom secara otomatis berdasarkan teks header)
    // Kita cek baris ke-2 (indeks 1) atau baris ke-1 (indeks 0)
    var headerRow = data[1]; 
    var idxKode = headerRow.indexOf("Kode Risiko");
    var idxKegiatan = headerRow.indexOf("Kegiatan/ Proses Bisnis");
    var idxPeristiwa = headerRow.indexOf("Risk Event/ Uraian Peristiwa Risiko");
    var idxPenyebab = headerRow.indexOf("Penyebab Risiko");
    var idxKemungkinan = headerRow.indexOf("Kemungkinan Inheren");
    var idxDampak = headerRow.indexOf("Dampak Inheren");
    var idxLevel = headerRow.indexOf("Level Risiko Inheren");
    
    // Fallback jika tidak sengaja terbaca di baris ke-1 (indeks 0) karena merge cells
    if (idxKode === -1) {
      headerRow = data[0];
      idxKode = headerRow.indexOf("Kode Risiko");
      idxKegiatan = headerRow.indexOf("Kegiatan/ Proses Bisnis");
    }
    if (idxPeristiwa === -1) idxPeristiwa = data[1].indexOf("Risk Event/ Peristiwa") !== -1 ? data[1].indexOf("Risk Event/ Peristiwa") : 5;
    if (idxPenyebab === -1) idxPenyebab = 6;
    if (idxKemungkinan === -1) idxKemungkinan = 11; // Posisi default standar Kolom L
    if (idxDampak === -1) idxDampak = 12;      // Posisi default standar Kolom M
    if (idxLevel === -1) idxLevel = 13;        // Posisi default standar Kolom N
    
    var belumDinilai = [];
    var sudahDinilai = [];
    
    // Mulai dari indeks 2 (Baris ke-3 Spreadsheet) untuk melewati 2 baris header bertingkat
    for (var i = 2; i < data.length; i++) {
      var row = data[i];
      if (!row || row.length === 0) continue;
      
      // Ambil data secara dinamis berdasarkan indeks yang sudah dicari di atas
      var kode = (idxKode < row.length) ? row[idxKode] : "";
      var kegiatan = (idxKegiatan < row.length) ? row[idxKegiatan] : "";
      var peristiwa = (idxPeristiwa < row.length) ? row[idxPeristiwa] : "";
      
      var kemungkinanInh = (idxKemungkinan < row.length) ? String(row[idxKemungkinan]).trim() : "";
      var dampakInh = (idxDampak < row.length) ? String(row[idxDampak]).trim() : "";
      var levelInh = (idxLevel < row.length) ? String(row[idxLevel]).trim() : "";
      
      // Lewati baris kosong jika tidak sengaja tersimpan baris hantu di sheet
      if (!kode && !kegiatan && !peristiwa) continue;
      
      var item = {
        rowNum: i + 1,
        kode: kode || "N/A",
        kegiatan: kegiatan || "-",
        peristiwa: peristiwa || "-",
        kemungkinan: kemungkinanInh,
        dampak: dampakInh,
        level: levelInh
      };
      
      // Pengelompokan: jika nilai kemungkinan kosong, strip (-), atau nol (0)
      if (kemungkinanInh === "" || kemungkinanInh === "-" || kemungkinanInh === "0") {
        belumDinilai.push(item);
      } else {
        sudahDinilai.push(item);
      }
    }
    
    return { belumDinilai: belumDinilai, sudahDinilai: sudahDinilai };
    
  } catch (e) {
    Logger.log("Error pada getDataUntukAnalisis: " + e.toString());
    return { belumDinilai: [], sudahDinilai: [] };
  }
}





// A. AMBIL SEMUA DATA DARI SHEET KONTEKS (BISA DIAKSES OPERATOR & ADMIN)
function getKonteksDataAll() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Konteks");
    
    if (!sheet) {
      Logger.log("Error: Sheet 'Konteks' tidak ditemukan.");
      return [];
    }
    
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return []; // Hanya ada baris header, return array kosong
    }
    
    var data = sheet.getDataRange().getValues();
    var list = [];
    
    // Membaca mulai dari baris ke-2 (indeks 1)
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      // Lewati jika baris benar-benar kosong
      if (!row || row.length < 2 || (!row[0] && !row[1])) continue;
      
      list.push({
        rowNum: i + 1, // Baris asli di spreadsheet untuk keperluan Edit
        no: row[0] !== undefined && row[0] !== "" ? row[0] : i,
        konteks: row[1] ? String(row[1]).trim() : "-",
        ket: row[2] ? String(row[2]).trim() : "-"
      });
    }
    return list;
  } catch (e) {
    Logger.log("Error pada getKonteksDataAll: " + e.toString());
    return [];
  }
}

// B. SIMPAN / TAMBAH DATA KONTEKS BARU (ADMIN ONLY)
function simpanKonteksBaru(formData, token) {
  try {
    pastikanAkses_(token, ["admin"]);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Konteks");
    if (!sheet) return { success: false, message: "Sheet 'Konteks' tidak ditemukan!" };
    
    // Ambil nomor urut terakhir secara otomatis jika kolom "no" kosong
    var lastRow = sheet.getLastRow();
    var nextNo = lastRow; 
    
    sheet.appendRow([nextNo, formData.konteks, formData.ket]);
    return { success: true, message: "Data Konteks berhasil ditambahkan!" };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// C. UPDATE / EDIT DATA KONTEKS EXISTING (ADMIN ONLY)
function updateKonteksData(formData, token) {
  try {
    pastikanAkses_(token, ["admin"]);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Konteks");
    if (!sheet) return { success: false, message: "Sheet 'Konteks' tidak ditemukan!" };
    
    var rowNum = parseInt(formData.rowNum);
    
    // Update data pada baris bersangkutan (Kolom B dan Kolom C)
    sheet.getRange(rowNum, 2).setValue(formData.konteks);
    sheet.getRange(rowNum, 3).setValue(formData.ket);
    
    return { success: true, message: "Data Konteks berhasil diperbarui!" };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}
// ===================================================================
// D. LANJUTAN ANALISIS RISIKO: PENGENDALIAN SAAT INI & RISIKO RESIDUAL
// ===================================================================
// Kolom DaftarResiko (0-index): 14=Aktivitas,15=Atribut,16=PenilaianKelemahan,
// 17=SimpulanEfektivitas,18=KemungkinanResidual,19=DampakResidual,20=LevelResidual

// Ambil risiko yang sudah dinilai Inheren tapi belum diisi Pengendalian & Residual
function getRisikoBelumResidual() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DaftarResiko");
    if (!sheet) return { belum: [], sudah: [] };
    var lastRow = sheet.getLastRow();
    if (lastRow <= 2) return { belum: [], sudah: [] };

    var data = sheet.getDataRange().getValues();
    var belum = [], sudah = [];

    for (var i = 2; i < data.length; i++) {
      var row = data[i];
      if (!row || row.length === 0) continue;
      var kode = row[3], kegiatan = row[2], peristiwa = row[5];
      var levelInh = row.length > 13 ? String(row[13]).trim() : "";
      if (!kode && !kegiatan) continue;
      if (!levelInh || levelInh === "-") continue; // belum dinilai inheren, skip dari modul ini

      var kemungkinanRes = row.length > 18 ? String(row[18]).trim() : "";
      var item = {
        rowNum: i + 1,
        kode: kode || "N/A",
        kegiatan: kegiatan || "-",
        peristiwa: peristiwa || "-",
        levelInheren: levelInh,
        aktivitas: row.length > 14 ? row[14] : "",
        atribut: row.length > 15 ? row[15] : "",
        penilaianKelemahan: row.length > 16 ? row[16] : "",
        simpulanEfektivitas: row.length > 17 ? row[17] : "",
        kemungkinanResidual: kemungkinanRes,
        dampakResidual: row.length > 19 ? row[19] : "",
        levelResidual: row.length > 20 ? row[20] : ""
      };

      if (!kemungkinanRes || kemungkinanRes === "-" || kemungkinanRes === "0") {
        belum.push(item);
      } else {
        sudah.push(item);
      }
    }
    return { belum: belum, sudah: sudah };
  } catch (e) {
    Logger.log("Error pada getRisikoBelumResidual: " + e.toString());
    return { belum: [], sudah: [] };
  }
}

// Simpan penilaian Pengendalian Saat Ini & Risiko Residual
function simpanPengendalianResidual(rowNum, dataPengendalian, token) {
  try {
    pastikanAkses_(token, ["admin", "operator"]);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DaftarResiko");
    if (!sheet) return { success: false, message: "Sheet DaftarResiko tidak ditemukan!" };

    var hasil = hitungLevelRisiko(dataPengendalian.kemungkinanResidual, dataPengendalian.dampakResidual);

    sheet.getRange(rowNum, 15).setValue(dataPengendalian.aktivitas);            // O
    sheet.getRange(rowNum, 16).setValue(dataPengendalian.atribut);              // P
    sheet.getRange(rowNum, 17).setValue(dataPengendalian.penilaianKelemahan);   // Q
    sheet.getRange(rowNum, 18).setValue(dataPengendalian.simpulanEfektivitas);  // R
    sheet.getRange(rowNum, 19).setValue(dataPengendalian.kemungkinanResidual);  // S
    sheet.getRange(rowNum, 20).setValue(dataPengendalian.dampakResidual);       // T
    sheet.getRange(rowNum, 21).setValue(hasil.skor + " (" + hasil.tingkat + ")"); // U

    return { success: true, message: "Pengendalian & Risiko Residual berhasil disimpan!" };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// ===================================================================
// E. EVALUASI RISIKO — keputusan mitigasi & prioritas (kolom 22 / W)
// ===================================================================
function getDataUntukEvaluasi() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DaftarResiko");
    if (!sheet) return [];
    var lastRow = sheet.getLastRow();
    if (lastRow <= 2) return [];

    var data = sheet.getDataRange().getValues();
    var list = [];

    for (var i = 2; i < data.length; i++) {
      var row = data[i];
      if (!row || row.length === 0) continue;
      var kode = row[3], kegiatan = row[2];
      if (!kode && !kegiatan) continue;

      var levelResidualRaw = row.length > 20 ? String(row[20]).trim() : "";
      var levelInherenRaw = row.length > 13 ? String(row[13]).trim() : "";
      var levelDipakai = levelResidualRaw && levelResidualRaw !== "-" ? levelResidualRaw : levelInherenRaw;
      if (!levelDipakai || levelDipakai === "-") continue; // belum ada dasar untuk dievaluasi

      // Skor dipakai untuk pengurutan prioritas (ambil angka sebelum kurung)
      var skorMatch = levelDipakai.match(/^(\d+)/);
      var skor = skorMatch ? parseInt(skorMatch[1]) : 0;

      list.push({
        rowNum: i + 1,
        kode: kode || "N/A",
        kegiatan: kegiatan || "-",
        levelDipakai: levelDipakai,
        sumberLevel: (levelResidualRaw && levelResidualRaw !== "-") ? "Residual" : "Inheren",
        skor: skor,
        keputusan: row.length > 22 ? (row[22] || "") : ""
      });
    }

    // Urutkan dari skor tertinggi (prioritas utama) ke terendah
    list.sort(function(a, b) { return b.skor - a.skor; });
    return list;
  } catch (e) {
    Logger.log("Error pada getDataUntukEvaluasi: " + e.toString());
    return [];
  }
}

function simpanKeputusanEvaluasi(rowNum, keputusan, token) {
  try {
    pastikanAkses_(token, ["admin", "operator"]);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DaftarResiko");
    if (!sheet) return { success: false, message: "Sheet DaftarResiko tidak ditemukan!" };
    sheet.getRange(rowNum, 23).setValue(keputusan); // Kolom W
    return { success: true, message: "Keputusan evaluasi berhasil disimpan!" };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// ===================================================================
// F. PENANGANAN RISIKO (RTP — RENCANA TINDAK PENGENDALIAN)
// ===================================================================
// Sheet "RTP" kolom: A=No, B=KodeRisiko, C=OpsiMitigasi, D=RencanaTindak,
// E=PIC, F=TargetWaktu, G=Anggaran, H=Status, I=TanggalInput

function getRTPList() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("RTP");
    if (!sheet) return [];
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];
    var data = sheet.getDataRange().getValues();
    var list = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row || !row[1]) continue;
      list.push({
        rowNum: i + 1,
        kodeRisiko: row[1],
        opsiMitigasi: row[2],
        rencanaTindak: row[3],
        pic: row[4],
        targetWaktu: row[5] ? new Date(row[5]).toISOString().slice(0, 10) : "",
        anggaran: row[6],
        status: row[7] || "Belum Mulai"
      });
    }
    return list;
  } catch (e) {
    Logger.log("Error pada getRTPList: " + e.toString());
    return [];
  }
}

function simpanRTP(formData, token) {
  try {
    pastikanAkses_(token, ["admin", "operator"]);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("RTP");
    if (!sheet) return { success: false, message: "Sheet 'RTP' tidak ditemukan! Buat sheet baru bernama RTP terlebih dahulu." };

    var nextNo = sheet.getLastRow();
    sheet.appendRow([
      nextNo, formData.kodeRisiko, formData.opsiMitigasi, formData.rencanaTindak,
      formData.pic, formData.targetWaktu, formData.anggaran, "Belum Mulai", new Date()
    ]);
    return { success: true, message: "Rencana Tindak Pengendalian berhasil ditambahkan!" };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function updateStatusRTP(rowNum, status, token) {
  try {
    pastikanAkses_(token, ["admin", "operator"]);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("RTP");
    if (!sheet) return { success: false, message: "Sheet RTP tidak ditemukan!" };
    sheet.getRange(rowNum, 8).setValue(status); // Kolom H
    return { success: true, message: "Status RTP berhasil diperbarui!" };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function hapusRTP(rowNum, token) {
  try {
    pastikanAkses_(token, ["admin"]);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("RTP");
    if (!sheet) return { success: false, message: "Sheet RTP tidak ditemukan!" };
    sheet.deleteRow(rowNum);
    return { success: true, message: "RTP berhasil dihapus!" };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// ===================================================================
// G. PEMANTAUAN & REVIU (per item RTP)
// ===================================================================
// Sheet "Pemantauan" kolom: A=No, B=RtpRowNum, C=KodeRisiko, D=TanggalReviu,
// E=Progres, F=Kendala, G=StatusBaru, H=Catatan, I=Penginput, J=WaktuInput

function getPemantauanByRTP(rtpRowNum) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Pemantauan");
    if (!sheet) return [];
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];
    var data = sheet.getDataRange().getValues();
    var list = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row || String(row[1]) !== String(rtpRowNum)) continue;
      list.push({
        rowNum: i + 1,
        tanggalReviu: row[3] ? new Date(row[3]).toISOString().slice(0, 10) : "",
        progres: row[4],
        kendala: row[5],
        statusBaru: row[6],
        catatan: row[7],
        penginput: row[8]
      });
    }
    return list.reverse(); // Terbaru dulu
  } catch (e) {
    Logger.log("Error pada getPemantauanByRTP: " + e.toString());
    return [];
  }
}

function simpanPemantauan(formData, token) {
  try {
    var session = pastikanAkses_(token, ["admin", "operator"]);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetPemantauan = ss.getSheetByName("Pemantauan");
    if (!sheetPemantauan) return { success: false, message: "Sheet 'Pemantauan' tidak ditemukan! Buat sheet baru bernama Pemantauan terlebih dahulu." };

    var nextNo = sheetPemantauan.getLastRow();
    sheetPemantauan.appendRow([
      nextNo, formData.rtpRowNum, formData.kodeRisiko, formData.tanggalReviu,
      formData.progres, formData.kendala, formData.statusBaru, formData.catatan,
      session.username, new Date()
    ]);

    // Sinkronkan status terbaru ke sheet RTP juga
    if (formData.statusBaru) {
      var sheetRTP = ss.getSheetByName("RTP");
      if (sheetRTP) sheetRTP.getRange(parseInt(formData.rtpRowNum), 8).setValue(formData.statusBaru);
    }

    return { success: true, message: "Catatan pemantauan berhasil disimpan!" };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}
