/**
 * TEMAN AI PINTAR 3D - BACKEND GOOGLE APPS SCRIPT (Code.gs)
 * Membaca API Key Groq dari Helaian 'apikey' Sel A1
 * Menyimpan Memori Vektor Wajah (128-Descriptor) & Memori Perbualan
 */

const SHEET_ID = "1caFK4WZ7fKk32-j4CXS7t22JGA9K806nLjHDXW2DKOg";
const FOLDER_ID = "18u5mjkdNw1LTZ7jGZ6K23FTqbNLChwja";

// Jalankan fungsi ini sekali di editor Apps Script untuk cipta tab automatik jika belum wujud
function setupAwal() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  
  let sheetKey = ss.getSheetByName("apikey");
  if (!sheetKey) {
    sheetKey = ss.insertSheet("apikey");
    sheetKey.getRange("A1").setValue("MASUKKAN_API_KEY_GROQ_DISINI");
  }
  
  let sheetWajah = ss.getSheetByName("MemoriWajah");
  if (!sheetWajah) {
    sheetWajah = ss.insertSheet("MemoriWajah");
    sheetWajah.appendRow(["Masa Daftar", "Nama Pengguna", "Vektor Wajah (128-Float JSON)"]);
  }
  
  let sheetLog = ss.getSheetByName("LogMemori");
  if (!sheetLog) {
    sheetLog = ss.insertSheet("LogMemori");
    sheetLog.appendRow(["Masa", "Nama", "Emosi Wajah", "Mod Kamera", "Mesej Pengguna", "Jawapan AI"]);
  }
}

function ambilApiKeyGroq() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName("apikey");
  if (!sheet) {
    setupAwal();
    sheet = ss.getSheetByName("apikey");
  }
  const apiKey = String(sheet.getRange("A1").getValue() || "").replace(/[\s"']/g, "");
  if (!apiKey || apiKey === "MASUKKAN_API_KEY_GROQ_DISINI") {
    throw new Error("Kunci API Groq belum dimasukkan dalam tab 'apikey' sel A1.");
  }
  return apiKey;
}

function doGet(e) {
  return prosesPermintaan(e);
}

function doPost(e) {
  return prosesPermintaan(e);
}

function prosesPermintaan(e) {
  try {
    let params = {};
    if (e && e.postData && e.postData.contents) {
      params = JSON.parse(e.postData.contents);
    } else if (e && e.parameter) {
      params = e.parameter;
    }

    const action = params.action || "getConfig";

    // 0. Uji kunci API dalam Sheet (tab 'apikey' sel A1) terus ke Groq
    if (action === "testKey") {
      let apiKey = "";
      try { apiKey = ambilApiKeyGroq(); } catch (kErr) {
        return responJSON({ status: "error", alive: false, maskedKey: "Tiada", keyLength: 0, message: kErr.message });
      }
      const t0 = Date.now();
      const r = UrlFetchApp.fetch("https://api.groq.com/openai/v1/models", {
        method: "get", headers: { "Authorization": "Bearer " + apiKey }, muteHttpExceptions: true
      });
      const code = r.getResponseCode();
      let msg = "OK";
      if (code !== 200) {
        try { msg = JSON.parse(r.getContentText()).error.message; } catch (e2) { msg = "HTTP " + code; }
      }
      return responJSON({
        status: code === 200 ? "success" : "error",
        alive: code === 200,
        maskedKey: topengKunci(apiKey),
        keyLength: apiKey.length,
        latencyMs: Date.now() - t0,
        message: msg
      });
    }

    // 1. Ambil konfigurasi & senarai memori wajah yang pernah disimpan
    if (action === "getConfig") {
      const apiKey = ambilApiKeyGroq();
      const ss = SpreadsheetApp.openById(SHEET_ID);
      let sheetWajah = ss.getSheetByName("MemoriWajah");
      let senaraiWajah = [];

      if (sheetWajah && sheetWajah.getLastRow() > 1) {
        const data = sheetWajah.getRange(2, 1, sheetWajah.getLastRow() - 1, 3).getValues();
        data.forEach(row => {
          if (row[1] && row[2]) {
            try {
              senaraiWajah.push({
                name: String(row[1]),
                descriptors: JSON.parse(row[2])
              });
            } catch (err) {}
          }
        });
      }

      return responJSON({
        status: "success",
        maskedKey: topengKunci(apiKey),
        faces: senaraiWajah
      });
    }

    // 2. Simpan memori vektor wajah pengguna ke Google Sheet & salinan JSON ke Google Drive
    if (action === "saveFace") {
      const ss = SpreadsheetApp.openById(SHEET_ID);
      let sheetWajah = ss.getSheetByName("MemoriWajah");
      if (!sheetWajah) {
        setupAwal();
        sheetWajah = ss.getSheetByName("MemoriWajah");
      }

      const nama = params.name || "Pengguna";
      const descriptorsJson = JSON.stringify(params.descriptors || []);
      const masa = new Date();

      // Kemas kini jika nama sudah wujud atau tambah baris baharu
      const lastRow = sheetWajah.getLastRow();
      let dikemaskini = false;
      if (lastRow > 1) {
        const data = sheetWajah.getRange(2, 1, lastRow - 1, 3).getValues();
        for (let i = 0; i < data.length; i++) {
          if (String(data[i][1]).toLowerCase() === nama.toLowerCase()) {
            sheetWajah.getRange(i + 2, 1, 1, 3).setValues([[masa, nama, descriptorsJson]]);
            dikemaskini = true;
            break;
          }
        }
      }
      if (!dikemaskini) {
        sheetWajah.appendRow([masa, nama, descriptorsJson]);
      }

      // Simpan fail sandaran profil wajah ke dalam Google Drive Folder
      try {
        const folder = DriveApp.getFolderById(FOLDER_ID);
        const namaFail = "MemoriWajah_" + nama.replace(/[^a-zA-Z0-9]/g, "_") + ".json";
        const files = folder.getFilesByName(namaFail);
        const kandunganFail = JSON.stringify({
          name: nama,
          updatedAt: masa.toISOString(),
          descriptors: params.descriptors
        }, null, 2);

        if (files.hasNext()) {
          files.next().setContent(kandunganFail);
        } else {
          folder.createFile(namaFail, kandunganFail, MimeType.PLAIN_TEXT);
        }
      } catch (driveErr) {}

      return responJSON({ status: "success", message: "Memori wajah disimpan." });
    }

    // 3. Panggilan terus ke Groq melalui pelayan Apps Script & simpan log memori
    if (action === "chat") {
      const apiKey = ambilApiKeyGroq();
      const nama = params.name || "Sahabat";
      const emosi = params.emotion || "Neutral / Tenang";
      const modKamera = params.cameraMode || "Aktif";
      const mesej = params.message || "";
      let sejarah = params.history || [];
      if (typeof sejarah === "string") { try { sejarah = JSON.parse(sejarah); } catch (h) { sejarah = []; } }
      sejarah = (Array.isArray(sejarah) ? sejarah : []).filter(m => m && m.role && m.content);

      const systemPrompt = `Anda ialah sahabat karib (kawan baik) dan AI serba tahu kepada ${nama}.
Anda bercakap dalam Bahasa Melayu Malaysia yang sangat natural, santai, mesra, empati, dan bijaksana seperti manusia sebenar.
Maklumat deria semasa:
- Nama di hadapan anda: ${nama}
- Mod Kamera: ${modKamera}
- Riak / Emosi Wajah Semasa: ${emosi}

Peraturan Penting:
1. Jika kamera aktif, sesuaikan nada perbualan secara semula jadi dengan emosi wajah ${nama} tanpa menyebut secara kaku "kamera saya mengesan wajah anda". Jika kamera disembunyikan/ditutup (Mod Suara Sahaja), fokus sepenuhnya pada intonasi dan topik perbualan.
2. Anda serba tahu dalam semua bidang (sains, pendidikan, sukan, pembaikan, koding, agama, kehidupan harian).
3. Jawab dengan padat, hidup, dan mesra (sekitar 1 hingga 4 ayat sahaja) supaya sangat sedap didengar apabila dibacakan melalui suara (Text-to-Speech). Jangan gunakan simbol markdown seperti bintang (*) atau tanda pagar (#) kerana jawapan akan dibaca terus oleh enjin suara.`;

      const messagesPayload = [
        { role: "system", content: systemPrompt },
        ...sejarah.slice(-8),
        { role: "user", content: mesej }
      ];

      const urlGroq = "https://api.groq.com/openai/v1/chat/completions";
      const options = {
        method: "post",
        contentType: "application/json",
        headers: {
          "Authorization": "Bearer " + apiKey
        },
        payload: JSON.stringify({
          model: "llama-3.3-70b-versatile",
          messages: messagesPayload,
          temperature: 0.72,
          max_tokens: 350
        }),
        muteHttpExceptions: true
      };

      const res = UrlFetchApp.fetch(urlGroq, options);
      const json = JSON.parse(res.getContentText());

      if (json.error) {
        return responJSON({ status: "error", message: json.error.message });
      }

      const jawapanAI = json.choices[0].message.content.trim();

      // Rekod ke dalam helaian LogMemori
      try {
        const ss = SpreadsheetApp.openById(SHEET_ID);
        let sheetLog = ss.getSheetByName("LogMemori");
        if (sheetLog) {
          sheetLog.appendRow([new Date(), nama, emosi, modKamera, mesej, jawapanAI]);
        }
      } catch (logErr) {}

      return responJSON({
        status: "success",
        reply: jawapanAI
      });
    }

    return responJSON({ status: "error", message: "Tindakan tidak dikenali." });
  } catch (err) {
    return responJSON({ status: "error", message: err.toString() });
  }
}

function topengKunci(k) {
  if (!k) return "Tiada";
  return k.length > 10 ? k.slice(0, 4) + "••••" + k.slice(-4) : "••••";
}

function responJSON(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
