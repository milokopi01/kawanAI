/**
 * AURA TERJEMAH — BACKEND GOOGLE APPS SCRIPT (Code.gs)
 * Penterjemah suara masa nyata: cakap dalam Bahasa Melayu (atau bahasa lain),
 * diterjemah ke mana-mana bahasa dunia menggunakan Groq.
 * Membaca API Key Groq dari Helaian 'apikey' Sel A1 (model pilihan: Sel B1).
 * Merekod setiap terjemahan ke helaian 'LogTerjemahan'.
 */

const SHEET_ID = "1caFK4WZ7fKk32-j4CXS7t22JGA9K806nLjHDXW2DKOg";

// Jalankan fungsi ini sekali di editor Apps Script untuk cipta tab automatik jika belum wujud
function setupAwal() {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  let sheetKey = ss.getSheetByName("apikey");
  if (!sheetKey) {
    sheetKey = ss.insertSheet("apikey");
    sheetKey.getRange("A1").setValue("MASUKKAN_API_KEY_GROQ_DISINI");
  }

  let sheetLog = ss.getSheetByName("LogTerjemahan");
  if (!sheetLog) {
    sheetLog = ss.insertSheet("LogTerjemahan");
    sheetLog.appendRow(["Masa", "Dari Bahasa", "Ke Bahasa", "Gaya", "Teks Asal", "Terjemahan", "Sebutan", "Model"]);
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
 * Kod ini membaca senarai model yang BENAR-BENAR tersedia untuk kunci anda.
 * Mahu paksa model tertentu? Tulis nama model dalam tab 'apikey' sel B1.
 * ===================================================================== */
const MODEL_PILIHAN = [
  "llama-3.3-70b-versatile",
  "openai/gpt-oss-120b",
  "moonshotai/kimi-k2-instruct-0905",
  "meta-llama/llama-4-maverick-17b-128e-instruct",
  "meta-llama/llama-4-scout-17b-16e-instruct",
  "qwen/qwen3-32b",
  "openai/gpt-oss-20b",
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

function susunCalonModel(apiKey, segarkan) {
  const ada = senaraiModelGroq(apiKey, segarkan);
  const manual = modelManualDariSheet();
  const calon = [];
  if (manual && (!ada.length || ada.indexOf(manual) >= 0)) calon.push(manual);
  MODEL_PILIHAN.forEach(m => { if ((!ada.length || ada.indexOf(m) >= 0) && calon.indexOf(m) < 0) calon.push(m); });
  ada.forEach(m => { if (calon.indexOf(m) < 0 && !/compound/i.test(m)) calon.push(m); });
  return calon;
}

function bersihkanOutput(t) {
  return String(t || "")
    .replace(/<think>[\s\S]*?<\/think>/gi, "")
    .replace(/```(?:json)?/gi, "")
    .trim();
}

function panggilGroq(apiKey, messagesPayload, opsyen) {
  opsyen = opsyen || {};
  let calon = susunCalonModel(apiKey, false);
  let ralatAkhir = "Tiada model chat tersedia untuk kunci API ini.";
  let dahSegar = false;

  for (let i = 0; i < calon.length && i < 5; i++) {
    const model = calon[i];
    const body = {
      model: model,
      messages: messagesPayload,
      temperature: opsyen.temperature != null ? opsyen.temperature : 0.2,
      max_tokens: opsyen.max_tokens || 1200
    };
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
      const teks = bersihkanOutput(json.choices[0].message && json.choices[0].message.content);
      if (teks) return { reply: teks, model: model };
      ralatAkhir = "Model " + model + " memulangkan jawapan kosong.";
      continue;
    }

    ralatAkhir = (json.error && json.error.message) || ("HTTP " + res.getResponseCode());
    if (/model|decommission|not exist|not found|access|deprecat/i.test(ralatAkhir)) {
      if (!dahSegar) {
        dahSegar = true;
        const baru = susunCalonModel(apiKey, true).filter(m => m !== model);
        calon = calon.slice(0, i + 1).concat(baru.filter(m => calon.slice(0, i + 1).indexOf(m) < 0));
      }
      continue;
    }
    break; // had kadar / kunci salah -> berhenti
  }
  throw new Error(ralatAkhir);
}

/* =====================================================================
 * ENJIN TERJEMAHAN
 * ===================================================================== */
const GAYA_TERJEMAHAN = {
  biasa: "natural, neutral everyday spoken",
  formal: "formal and polite (suitable for business, officials or elders)",
  santai: "casual and friendly (like talking to a friend)"
};

function terjemahTeks(apiKey, teks, sumber, sasaran, gaya, romanize, konteks) {
  const register = GAYA_TERJEMAHAN[gaya] || GAYA_TERJEMAHAN.biasa;

  const systemPrompt = `You are a professional real-time speech interpreter.
Translate every user message from ${sumber} into ${sasaran}.

Rules:
1. The input comes from speech-to-text, so it may lack punctuation, contain slang, Malaysian colloquial words (e.g. "tak", "nak", "dah", "lah", "kot", "ke", "je", "camne", "boleh tak"), mixed English words, or small recognition mistakes. Understand the intended meaning first, then translate it faithfully.
2. Use a ${register} register in ${sasaran}. The result will be read aloud by a text-to-speech voice, so write it as natural speech with correct punctuation.
3. Keep names, numbers, dates, prices and places accurate. Never add explanations, notes, alternatives or quotation marks.
4. If the message is a question, keep it a question. If it is already in ${sasaran}, just return it cleaned up.
5. Reply ONLY with one JSON object and nothing else:
{"translation": "<the ${sasaran} translation>", "romanization": "<${romanize
    ? `Latin-alphabet pronunciation guide of the translation ONLY if ${sasaran} uses a non-Latin script (e.g. pinyin with tones, romaji, romanized Arabic/Thai/Hindi/Korean). Otherwise an empty string`
    : "always an empty string"}>"}`;

  const messages = [{ role: "system", content: systemPrompt }];
  (konteks || []).slice(-4).forEach(k => {
    if (k && k.src && k.tgt) {
      messages.push({ role: "user", content: String(k.src) });
      messages.push({ role: "assistant", content: JSON.stringify({ translation: String(k.tgt), romanization: "" }) });
    }
  });
  messages.push({ role: "user", content: teks });

  const hasil = panggilGroq(apiKey, messages, { temperature: 0.2, max_tokens: 1200 });

  let terjemahan = "", sebutan = "";
  const padanan = hasil.reply.match(/\{[\s\S]*\}/);
  if (padanan) {
    try {
      const obj = JSON.parse(padanan[0]);
      terjemahan = String(obj.translation || "").trim();
      sebutan = String(obj.romanization || "").trim();
    } catch (e) {}
  }
  if (!terjemahan) terjemahan = hasil.reply.replace(/^["']|["']$/g, "").trim();
  if (sebutan && sebutan.toLowerCase() === terjemahan.toLowerCase()) sebutan = "";

  return { translation: terjemahan, romanization: sebutan, model: hasil.model };
}

/* =====================================================================
 * PINTU MASUK WEB APP
 * ===================================================================== */
function doGet(e) { return prosesPermintaan(e); }
function doPost(e) { return prosesPermintaan(e); }

function prosesPermintaan(e) {
  try {
    let params = {};
    if (e && e.postData && e.postData.contents) {
      params = JSON.parse(e.postData.contents);
    } else if (e && e.parameter) {
      params = e.parameter;
    }

    const action = params.action || "getConfig";

    // 0. Uji kunci API dalam Sheet (tab 'apikey' sel A1)
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

    // 1. Semak sambungan
    if (action === "getConfig") {
      const apiKey = ambilApiKeyGroq();
      return responJSON({ status: "success", maskedKey: topengKunci(apiKey) });
    }

    // 2. Terjemah teks / ucapan
    if (action === "translate") {
      const apiKey = ambilApiKeyGroq();
      const teks = String(params.text || "").trim();
      if (!teks) return responJSON({ status: "error", message: "Tiada teks untuk diterjemah." });

      const sumber = String(params.source || "Malay");
      const sasaran = String(params.target || "English");
      const gaya = String(params.style || "biasa");
      const romanize = params.romanize === true || params.romanize === "true";
      let konteks = params.context || [];
      if (typeof konteks === "string") { try { konteks = JSON.parse(konteks); } catch (c) { konteks = []; } }
      if (!Array.isArray(konteks)) konteks = [];

      let hasil;
      try {
        hasil = terjemahTeks(apiKey, teks, sumber, sasaran, gaya, romanize, konteks);
      } catch (gErr) {
        return responJSON({ status: "error", message: gErr.message });
      }

      try {
        const ss = SpreadsheetApp.openById(SHEET_ID);
        let sheetLog = ss.getSheetByName("LogTerjemahan");
        if (!sheetLog) { setupAwal(); sheetLog = ss.getSheetByName("LogTerjemahan"); }
        sheetLog.appendRow([new Date(), sumber, sasaran, gaya, teks, hasil.translation, hasil.romanization, hasil.model]);
      } catch (logErr) {}

      return responJSON({
        status: "success",
        translation: hasil.translation,
        romanization: hasil.romanization,
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
