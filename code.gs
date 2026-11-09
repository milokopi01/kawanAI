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

/* =====================================================================
 * PEMILIHAN MODEL GROQ AUTOMATIK
 * Groq kerap menukar / menamatkan model. Kod ini membaca senarai model
 * yang BENAR-BENAR tersedia untuk kunci anda, dan memilih yang terbaik.
 * Mahu paksa model tertentu? Tulis nama model dalam tab 'apikey' sel B1.
 * ===================================================================== */
const MODEL_PILIHAN = [
  "openai/gpt-oss-120b",
  "llama-3.3-70b-versatile",
  "meta-llama/llama-4-maverick-17b-128e-instruct",
  "moonshotai/kimi-k2-instruct-0905",
  "meta-llama/llama-4-scout-17b-16e-instruct",
  "openai/gpt-oss-20b",
  "qwen/qwen3-32b",
  "llama-3.1-8b-instant"
];

function senaraiModelGroq(apiKey, segarkan) {
  const cache = CacheService.getScriptCache();
  if (!segarkan) {
    const c = cache.get("groq_models_v2");
    if (c) { try { return JSON.parse(c); } catch (e) {} }
  }
  const r = UrlFetchApp.fetch("https://api.groq.com/openai/v1/models", {
    method: "get", headers: { "Authorization": "Bearer " + apiKey }, muteHttpExceptions: true
  });
  if (r.getResponseCode() !== 200) return [];
  let ids = [];
  try {
    ids = (JSON.parse(r.getContentText()).data || [])
      .filter(m => m && m.id && m.active !== false)
      .map(m => String(m.id))
      .filter(id => !/whisper|guard|tts|playai|orpheus|prompt-guard|safeguard|distil/i.test(id));
  } catch (e) {}
  if (ids.length) cache.put("groq_models_v2", JSON.stringify(ids), 21600);
  return ids;
}

function modelManualDariSheet() {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName("apikey");
    if (!sheet) return "";
    return String(sheet.getRange("B1").getValue() || "").trim();
  } catch (e) { return ""; }
}

// Pulangkan senarai model untuk dicuba mengikut keutamaan
function susunCalonModel(apiKey, segarkan) {
  const ada = senaraiModelGroq(apiKey, segarkan);
  const manual = modelManualDariSheet();
  const calon = [];
  if (manual && (!ada.length || ada.indexOf(manual) >= 0)) calon.push(manual);
  MODEL_PILIHAN.forEach(m => { if ((!ada.length || ada.indexOf(m) >= 0) && calon.indexOf(m) < 0) calon.push(m); });
  ada.forEach(m => { if (calon.indexOf(m) < 0 && !/compound/i.test(m)) calon.push(m); });
  return calon;
}

function bersihkanJawapan(t) {
  return String(t || "")
    .replace(/<think>[\s\S]*?<\/think>/gi, "")
    .replace(/[*#`]/g, "")
    .trim();
}

function panggilGroqChat(apiKey, messagesPayload) {
  let calon = susunCalonModel(apiKey, false);
  let ralatAkhir = "Tiada model chat tersedia untuk kunci API ini.";
  let dahSegar = false;

  for (let i = 0; i < calon.length && i < 5; i++) {
    const model = calon[i];
    const body = { model: model, messages: messagesPayload, temperature: 0.72, max_tokens: 700 };
    if (/^openai\/gpt-oss/i.test(model)) body.reasoning_effort = "low";
    if (/qwen3/i.test(model)) body.reasoning_format = "hidden";

    const res = UrlFetchApp.fetch("https://api.groq.com/openai/v1/chat/completions", {
      method: "post",
      contentType: "application/json",
      headers: { "Authorization": "Bearer " + apiKey },
      payload: JSON.stringify(body),
      muteHttpExceptions: true
    });
    let json = {};
    try { json = JSON.parse(res.getContentText()); } catch (e) { json = { error: { message: "HTTP " + res.getResponseCode() } }; }

    if (!json.error && json.choices && json.choices[0]) {
      const teks = bersihkanJawapan(json.choices[0].message && json.choices[0].message.content);
      if (teks) return { reply: teks, model: model };
      ralatAkhir = "Model " + model + " memulangkan jawapan kosong.";
      continue;
    }

    ralatAkhir = (json.error && json.error.message) || ("HTTP " + res.getResponseCode());
    // Ralat berkaitan model -> cuba model seterusnya (segarkan senarai sekali)
    if (/model|decommission|not exist|not found|access|deprecat/i.test(ralatAkhir)) {
      if (!dahSegar) {
        dahSegar = true;
        const baru = susunCalonModel(apiKey, true).filter(m => m !== model);
        calon = calon.slice(0, i + 1).concat(baru.filter(m => calon.slice(0, i + 1).indexOf(m) < 0));
      }
      continue;
    }
    // Ralat lain (had kadar, kunci salah) -> berhenti
    break;
  }
  throw new Error(ralatAkhir);
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
        model: code === 200 ? (susunCalonModel(apiKey, true)[0] || "Tiada") : "",
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
3. Jawab dengan padat, hidup, dan mesra (sekitar 1 hingga 4 ayat sahaja) supaya sangat sedap didengar apabila dibacakan melalui suara (Text-to-Speech). Jawapan mesti teks biasa sahaja kerana ia akan dibacakan dengan suara perempuan Melayu. Jangan gunakan emoji atau simbol markdown seperti bintang (*) atau tanda pagar (#) kerana jawapan akan dibaca terus oleh enjin suara.`;

      const messagesPayload = [
        { role: "system", content: systemPrompt },
        ...sejarah.slice(-8),
        { role: "user", content: mesej }
      ];

      let hasil;
      try {
        hasil = panggilGroqChat(apiKey, messagesPayload);
      } catch (gErr) {
        return responJSON({ status: "error", message: gErr.message });
      }
      const jawapanAI = hasil.reply;

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
        reply: jawapanAI,
        model: hasil.model
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
