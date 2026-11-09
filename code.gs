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

function ambilApiKeyGroq(segar) {
  // Cache 10 minit supaya tidak buka Google Sheet setiap kali berbual (lebih laju)
  const cache = CacheService.getScriptCache();
  if (!segar) {
    const k = cache.get("groq_key_v1");
    if (k) return k;
  }
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
  cache.put("groq_key_v1", apiKey, 600);
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
  const cache = CacheService.getScriptCache();
  const c = cache.get("groq_model_b1");
  if (c !== null) return c === "-" ? "" : c;
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName("apikey");
    const v = sheet ? String(sheet.getRange("B1").getValue() || "").trim() : "";
    cache.put("groq_model_b1", v || "-", 600);
    return v;
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

function bersihkanJawapan(t, mentah) {
  const x = String(t || "").replace(/<think>[\s\S]*?<\/think>/gi, "").replace(/[*#`]/g, "").trim();
  if (mentah) return x;
  return bahasaMelayukan(String(t || "")
    .replace(/<think>[\s\S]*?<\/think>/gi, "")
    .replace(/[*#`]/g, "")
    .trim());
}

/* =====================================================================
 * PENAPIS BAHASA: tukar perkataan Bahasa Indonesia -> Bahasa Melayu Malaysia
 * (jaring keselamatan jika model AI tergelincir ke Bahasa Indonesia)
 * ===================================================================== */
const KAMUS_ID_KE_MS = [
  ["nggak", "tak"], ["enggak", "tak"], ["ngga", "tak"], ["gak", "tak"], ["gk", "tak"],
  ["bisa", "boleh"], ["banget", "sangat"], ["gimana", "macam mana"], ["bagaimana kabarmu", "apa khabar"],
  ["kabarmu", "khabar awak"], ["kabar", "khabar"], ["kamu", "awak"], ["kalian", "korang"], ["gue", "saya"], ["gua", "saya"],
  ["lu", "awak"], ["mau", "nak"], ["pengen", "nak"], ["pingin", "nak"], ["udah", "dah"], ["aja", "je"], ["saja kok", "je"],
  ["ayo", "jom"], ["yuk", "jom"], ["kayak", "macam"], ["seperti apa", "macam mana"], ["gitu", "macam tu"], ["gini", "macam ni"],
  ["emang", "memang"], ["bener", "betul"], ["beneran", "betul ke"], ["cewek", "perempuan"], ["cowok", "lelaki"],
  ["teman", "kawan"], ["silakan", "sila"], ["berbicara", "bercakap"], ["bicara", "cakap"], ["ngomong", "cakap"],
  ["mengerti", "faham"], ["ngerti", "faham"], ["paham", "faham"], ["memahami", "memahami"], ["coba", "cuba"], ["mencoba", "mencuba"],
  ["kapan", "bila"], ["harus", "perlu"], ["lucu", "kelakar"], ["keren", "hebat"], ["mobil", "kereta"], ["uang", "duit"],
  ["kantor", "pejabat"], ["sepeda", "basikal"], ["rumah sakit", "hospital"], ["kamar", "bilik"], ["gratis", "percuma"],
  ["handphone", "telefon"], ["hp", "telefon"], ["ponsel", "telefon bimbit"], ["libur", "cuti"], ["liburan", "percutian"],
  ["karena", "kerana"], ["soalnya", "sebab"], ["sekarang ini", "sekarang"], ["terus terang", "terus terang"],
  ["nih", "ni"], ["tuh", "tu"], ["dong", ""], ["deh", ""], ["sih", ""], ["kok", ""], ["lho", ""], ["loh", ""],
  ["sampai jumpa", "jumpa lagi"], ["selamat tinggal", "selamat tinggal"], ["apa kabar", "apa khabar"],
  ["semangat ya", "semangat ya"], ["ketemu", "jumpa"], ["bertemu", "berjumpa"], ["nanti malam", "malam nanti"],
  ["tolong", "tolong"], ["pakai", "guna"], ["memakai", "menggunakan"], ["gampang", "senang"], ["mudah-mudahan", "harap-harap"],
  ["semoga", "semoga"], ["bagus banget", "bagus sangat"], ["kesal", "geram"], ["capek", "penat"], ["capai", "capai"],
  ["lagi apa", "tengah buat apa"], ["sedang apa", "tengah buat apa"], ["dulu ya", "dulu ya"]
];
function bahasaMelayukan(teks) {
  let t = String(teks || "");
  // Susun frasa panjang dahulu
  const kamus = KAMUS_ID_KE_MS.slice().sort((a, b) => b[0].length - a[0].length);
  kamus.forEach(([id, ms]) => {
    if (id === ms) return;
    const re = new RegExp("(^|[^A-Za-zÀ-ÿ])(" + id.replace(/[.*+?^${}()|[\]\\]/g, "\\$&") + ")(?=$|[^A-Za-zÀ-ÿ])", "gi");
    t = t.replace(re, (m, pre, kata) => {
      let ganti = ms;
      if (ganti && kata[0] === kata[0].toUpperCase() && kata[0] !== kata[0].toLowerCase()) {
        ganti = ganti[0].toUpperCase() + ganti.slice(1);
      }
      return pre + ganti;
    });
  });
  // -itas -> -iti (aktivitas -> aktiviti, kualitas -> kualiti)
  t = t.replace(/\b([A-Za-z]{3,})itas\b/g, "$1iti");
  return t.replace(/\s{2,}/g, " ").replace(/\s+([,.!?])/g, "$1").trim();
}

function panggilGroqChat(apiKey, messagesPayload, mentah) {
  let calon = susunCalonModel(apiKey, false);
  let ralatAkhir = "Tiada model chat tersedia untuk kunci API ini.";
  let dahSegar = false;

  // Model yang pernah gagal (tiada akses) diingat 6 jam & dilangkau terus -> lebih laju
  const cache = CacheService.getScriptCache();
  let disekat = [];
  try { disekat = JSON.parse(cache.get("groq_model_disekat") || "[]"); } catch (e) {}
  const bolehGuna = calon.filter(m => disekat.indexOf(m) < 0);
  if (bolehGuna.length) calon = bolehGuna;

  for (let i = 0; i < calon.length && i < 10; i++) {
    const model = calon[i];
    const body = { model: model, messages: messagesPayload, temperature: 0.55, max_tokens: 350 };
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
      const teks = bersihkanJawapan(json.choices[0].message && json.choices[0].message.content, mentah);
      if (teks) return { reply: teks, model: model };
      ralatAkhir = "Model " + model + " memulangkan jawapan kosong.";
      continue;
    }

    ralatAkhir = (json.error && json.error.message) || ("HTTP " + res.getResponseCode());
    // Ralat berkaitan model -> cuba model seterusnya (segarkan senarai sekali)
    if (/model|decommission|not exist|not found|access|deprecat|permission|blocked/i.test(ralatAkhir)) {
      if (disekat.indexOf(model) < 0) { disekat.push(model); cache.put("groq_model_disekat", JSON.stringify(disekat), 21600); }
      if (!dahSegar) {
        dahSegar = true;
        const baru = susunCalonModel(apiKey, true).filter(m => m !== model && disekat.indexOf(m) < 0);
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
      try { apiKey = ambilApiKeyGroq(true); } catch (kErr) {
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

    // 2b. Mod terjemahan
    if (action === "transcribe") {
      const apiKeyT = ambilApiKeyGroq(false);
      if (!apiKeyT) return responJSON({ success: false, error: "Tiada API key" });
      const mime = String(params.mime || "audio/webm").split(";")[0];
      const ext = /mp4|m4a|aac/.test(mime) ? "m4a" : /ogg/.test(mime) ? "ogg" : /wav/.test(mime) ? "wav" : "webm";
      const blob = Utilities.newBlob(Utilities.base64Decode(String(params.audio || "")), mime, "suara." + ext);
      const petunjuk = "Perbualan santai dalam Bahasa Melayu Malaysia, kadang-kadang bercampur Bahasa Inggeris. Nama: "
        + String(params.aiName || "Aura") + ", " + String(params.name || "") + ". Contoh: awak, saya, tak, nak, dah, je, boleh, macam mana, okey.";
      let ralatT = "";
      const modelW = ["whisper-large-v3", "whisper-large-v3-turbo"];
      for (let i = 0; i < modelW.length; i++) {
        const r = UrlFetchApp.fetch("https://api.groq.com/openai/v1/audio/transcriptions", {
          method: "post", headers: { "Authorization": "Bearer " + apiKeyT }, muteHttpExceptions: true,
          payload: { file: blob, model: modelW[i], language: "ms", prompt: petunjuk, temperature: "0", response_format: "json" }
        });
        let j = {}; try { j = JSON.parse(r.getContentText()); } catch (x) {}
        if (j && typeof j.text === "string") return responJSON({ success: true, text: j.text.trim(), model: modelW[i] });
        ralatT = (j && j.error && j.error.message) || ("HTTP " + r.getResponseCode());
      }
      return responJSON({ success: false, error: ralatT });
    }

    if (action === "translate") {
      const apiKey = ambilApiKeyGroq();
      const sasaran = String(params.target || "English");
      const prompt = `Anda penterjemah profesional. Terjemahkan teks pengguna ke dalam ${sasaran}. Jika teks sudah dalam ${sasaran}, terjemahkan ke Bahasa Melayu Malaysia (bukan Bahasa Indonesia). Teks datang daripada pengecaman suara, jadi betulkan perkataan tersalah dengar ikut konteks. Pulangkan HANYA hasil terjemahan, tanpa penjelasan, tanpa tanda petik, tanpa emoji.`;
      let hasilT;
      try {
        hasilT = panggilGroqChat(apiKey, [{ role: "system", content: prompt }, { role: "user", content: String(params.message || "") }], true);
      } catch (tErr) { return responJSON({ status: "error", message: tErr.message }); }
      let r = hasilT.reply;
      if (/melayu/i.test(sasaran)) r = bahasaMelayukan(r);
      return responJSON({ status: "success", reply: r, model: hasilT.model });
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

      const namaAI = String(params.aiName || "Aura").trim().slice(0, 40) || "Aura";
      const systemPrompt = `Nama anda ialah ${namaAI}. Anda sahabat karib dan AI serba tahu kepada ${nama}. Jika ${nama} tanya nama anda, jawab "${namaAI}". Jika ${nama} panggil "${namaAI}", itu bermaksud dia bercakap dengan anda.

NOTA: Mesej ${nama} biasanya datang daripada pengecaman suara. Jika ada perkataan yang pelik atau tersalah dengar, fahami maksud sebenar ikut konteks perbualan dan bunyi yang hampir sama. Jangan tegur kesilapan ejaan itu.

BAHASA (PALING PENTING):
Anda WAJIB bercakap dalam Bahasa Melayu MALAYSIA sahaja, gaya santai orang Malaysia bersembang. JANGAN SEKALI-KALI guna Bahasa Indonesia.
Gunakan: awak / kau, saya / aku, tak, nak, dah, je, boleh, macam mana, betul, faham, cuba, perlu, kereta, duit, telefon, kawan, bila, sebab, kerana, sangat, jom, lah, kan, ke, tu, ni.
DILARANG (ini Bahasa Indonesia): kamu, bisa, nggak, gak, banget, gimana, mau, udah, aja, ayo, kayak, emang, bener, ngomong, ngerti, paham, coba, kapan, harus, mobil, uang, kantor, karena, dong, deh, sih, kok, nih, tuh, keren, capek, pakai, ketemu, liburan.
Ejaan Malaysia: "faham" bukan "paham", "kerana" bukan "karena", "khabar" bukan "kabar", "aktiviti" bukan "aktivitas", "universiti" bukan "universitas".

MAKLUMAT SEMASA:
- Nama pengguna: ${nama}
- Mod kamera: ${modKamera}
- Emosi wajah: ${emosi}

PERATURAN:
1. Mesej pengguna datang daripada pengecaman suara, jadi mungkin ada perkataan yang tersalah dengar. Teka maksud sebenar ${nama} berdasarkan konteks perbualan dan jawab apa yang dia betul-betul maksudkan. Jika langsung tak jelas, tanya semula secara ringkas.
2. Jawab TERUS kepada soalan atau perkara terkini yang ${nama} sebut. Jangan ulang jawapan lama dan jangan lari topik.
3. Sesuaikan nada dengan emosi wajah secara semula jadi, tanpa sebut "kamera saya mengesan".
4. Anda serba tahu (sains, pendidikan, sukan, pembaikan, koding, agama, kehidupan harian).
5. Jawab pendek dan padat, 1 hingga 3 ayat sahaja, kerana jawapan akan dibacakan dengan suara. Teks biasa sahaja, tiada emoji, tiada simbol seperti * atau #.`;

      const messagesPayload = [
        { role: "system", content: systemPrompt },
        ...sejarah.slice(-8).map(m => ({ role: m.role === "user" ? "user" : "assistant", content: m.role === "user" ? String(m.content) : bahasaMelayukan(m.content) })),
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
