/**
 * Klubbsalg: Google Sheet-bakend (Apps Script)
 *
 * OPPSETT (ca. 5 minutter)
 * 1. Lag et nytt Google Sheet, for eksempel «Klubbsalg».
 * 2. Åpne Utvidelser > Apps Script. Slett innholdet i Code.gs og lim inn hele denne filen.
 * 3. Åpne Prosjektinnstillinger (tannhjulet) > Skriptegenskaper > Legg til skriptegenskap:
 *      Egenskap: ADMIN_PASSWORD     Verdi: passordet du vil bruke for å slette
 *    Uten dette passordet er all sletting låst.
 * 4. Trykk Distribuer > Ny distribusjon > velg type «Nettapp»:
 *      Kjør som: Meg        Hvem har tilgang: Alle
 *    Godkjenn tilgangen når Google spør, og kopier «Nettapp-URL» (slutter på /exec).
 * 5. Lim URL-en inn i Klubbsalg-appen under «Google Sheet», og trykk Koble til.
 *
 * Etter endringer i denne koden: Distribuer > Administrer distribusjoner > rediger > Ny versjon.
 *
 * SIKKERHET
 * - Alle som har nettapp-URL-en kan registrere salg og lese oversikten. Del den bare med selgerne.
 * - Sletting, prisendring og fjerning av selger/produkt krever ADMIN_PASSWORD. Passordet sjekkes
 *   her på serveren og ligger aldri i appen.
 * - Selgeren kan endre eller slette sine egne salg i 30 minutter uten passord. Salget kjennes igjen på
 *   en hemmelig kode som ligger på selgerens telefon (i arket lagres bare en sjekksum av den), og
 *   tidsgrensen regnes av serveren. Vil du ha en annen grense, legg til skriptegenskapen
 *   EDIT_WINDOW_MINUTES (for eksempel 15). Alle endringer logges i kolonnene «Endret» og «Endringer».
 * - Etter 5 feil passord er alle slike handlinger låst i 10 minutter.
 * - Sletting er myk: raden blir stående i arket med tidspunkt i kolonnen «Slettet».
 *   Tøm den cellen for å gjenopprette salget.
 */

var SHEET_SALES = "Salg";
var SHEET_PRODUCTS = "Produkter";
var SHEET_SELLERS = "Selgere";
var SHEET_STANDS = "Stand";
var SHEET_STOCK = "Lager";
var SALES_HEADERS = ["ID", "Tid", "Selger", "Produkt", "Antall", "Normalpris per stk", "Normal totalt",
  "Solgt totalt", "Rabatt", "Betaling", "Kommentar", "Slettet", "Slettet av", "TS", "Mottatt", "Enhet", "Endret", "Endringer", "Stand", "Størrelse", "Valuta", "Kurs", "Beløp lokal valuta", "Ordre"];
var DEFAULT_EDIT_MINUTES = 30;
var PAYMENTS = ["Vipps", "Kontant", "Bank", "Annet"];
var TZ = "Europe/Oslo";
var MAX_FAILS = 5;
var LOCK_SECONDS = 600;
var MAX_PUSH = 200;

/* ---------- Inngang ---------- */

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  try {
    var action = (e && e.parameter && e.parameter.action) || "all";
    if (action === "ping") return json_({ ok: true, app: "klubbsalg", version: 2 });
    if (action === "all") return json_(readAll_());
    return json_({ ok: false, error: "ukjent_handling" });
  } catch (err) {
    return json_({ ok: false, error: "serverfeil", detail: String(err) });
  }
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  var locked = false;
  try {
    lock.waitLock(20000);
    locked = true;
    var body = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    switch (body.action) {
      case "push": return json_(push_(body));
      case "ownDelete": return json_(ownDelete_(body));
      case "ownEdit": return json_(ownEdit_(body));
      case "checkPassword": return json_(withPassword_(body, function () { return { ok: true }; }));
      case "delete": return json_(withPassword_(body, function () { return deleteSales_(body); }));
      case "deleteAll": return json_(withPassword_(body, function () { return deleteAll_(body); }));
      case "setPrice": return json_(withPassword_(body, function () { return setPrice_(body); }));
      case "setSizes": return json_(withPassword_(body, function () { return setPrice_(body); }));
      case "setImage": return json_(withPassword_(body, function () { return setImage_(body); }));
      case "setStand": return json_(withPassword_(body, function () { return setStand_(body); }));
      case "removeStand": return json_(withPassword_(body, function () { return removeName_(SHEET_STANDS, body.name); }));
      case "setStock": return json_(withPassword_(body, function () { return setStock_(body); }));
      case "removeProduct": return json_(withPassword_(body, function () { return removeName_(SHEET_PRODUCTS, body.name); }));
      case "removeSeller": return json_(withPassword_(body, function () { return removeName_(SHEET_SELLERS, body.name); }));
      default: return json_({ ok: false, error: "ukjent_handling" });
    }
  } catch (err) {
    return json_({ ok: false, error: "serverfeil", detail: String(err) });
  } finally {
    if (locked) { try { lock.releaseLock(); } catch (x) {} }
  }
}

/* ---------- Passord ---------- */

function withPassword_(body, fn) {
  var pw = PropertiesService.getScriptProperties().getProperty("ADMIN_PASSWORD");
  if (!pw) return { ok: false, error: "passord_ikke_satt" };
  var cache = CacheService.getScriptCache();
  var fails = Number(cache.get("pw_fails") || 0);
  if (fails >= MAX_FAILS) return { ok: false, error: "last", retryAfterSec: LOCK_SECONDS };
  if (String(body.password || "") !== pw) {
    cache.put("pw_fails", String(fails + 1), LOCK_SECONDS);
    return { ok: false, error: "feil_passord", attemptsLeft: MAX_FAILS - fails - 1 };
  }
  cache.remove("pw_fails");
  return fn();
}

/* ---------- Ark ---------- */

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }

function sheet_(name, headers) {
  var s = ss_().getSheetByName(name);
  if (!s) {
    s = ss_().insertSheet(name);
    s.getRange(1, 1, 1, headers.length).setValues([headers]);
    s.setFrozenRows(1);
  } else if (s.getMaxColumns() < headers.length) {
    s.insertColumnsAfter(s.getMaxColumns(), headers.length - s.getMaxColumns());
    s.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else if (headers.length > 1 && s.getRange(1, headers.length, 1, 1).getValues()[0][0] === "") {
    s.getRange(1, 1, 1, headers.length).setValues([headers]);   // eldre ark: legg til nye kolonneoverskrifter
  }
  return s;
}
function salesSheet_() { return sheet_(SHEET_SALES, SALES_HEADERS); }
function productsSheet_() { return sheet_(SHEET_PRODUCTS, ["Navn", "Standardpris", "Innkjøpspris", "Størrelser", "Bilde"]); }
function standsSheet_() { return sheet_(SHEET_STANDS, ["Navn", "Valuta", "Kurs til NOK"]); }
function stockSheet_() { return sheet_(SHEET_STOCK, ["Produkt", "Størrelse", "Antall", "Talt (ts)", "Talt"]); }
function sellersSheet_() { return sheet_(SHEET_SELLERS, ["Navn"]); }

function txt_(v, max) { return String(v == null ? "" : v).replace(/\s+/g, " ").trim().slice(0, max); }
function num_(v) { var n = Number(v); return isFinite(n) ? n : 0; }

function cleanSale_(x) {
  if (!x || typeof x !== "object") return null;
  var id = txt_(x.id, 64), seller = txt_(x.seller, 100), product = txt_(x.product, 100);
  if (!id || !seller || !product) return null;
  var qty = Math.round(num_(x.qty));
  if (qty < 1 || qty > 100000) return null;
  var normalPrice = Math.max(0, num_(x.normalPrice));
  var soldTotal = num_(x.soldTotal);
  if (soldTotal < 0 || soldTotal > 100000000) return null;
  var normalTotal = x.normalTotal == null ? (normalPrice > 0 ? qty * normalPrice : soldTotal) : num_(x.normalTotal);
  var discount = x.discount == null ? normalTotal - soldTotal : num_(x.discount);
  var payment = PAYMENTS.indexOf(txt_(x.payment, 20)) >= 0 ? txt_(x.payment, 20) : "Annet";
  var ts = num_(x.ts) > 0 ? num_(x.ts) : Date.now();
  var cur = txt_(x.cur, 8).toUpperCase() || "NOK";
  var rate = num_(x.rate) > 0 ? num_(x.rate) : 1;
  if (cur === "NOK") rate = 1;
  var local = x.local == null || x.local === "" ? (rate === 1 ? soldTotal : Math.round(soldTotal / rate * 100) / 100) : Math.max(0, num_(x.local));
  return { id: id, ts: ts, seller: seller, product: product, qty: qty, normalPrice: normalPrice,
    normalTotal: normalTotal, soldTotal: soldTotal, discount: discount, payment: payment, note: txt_(x.note, 500),
    stand: txt_(x.stand, 60), size: txt_(x.size, 30), cur: cur, rate: rate, local: local, order: txt_(x.order, 64) };
}

function rowOf_(c, deviceHash) {
  return [c.id, Utilities.formatDate(new Date(c.ts), TZ, "yyyy-MM-dd HH:mm"), c.seller, c.product, c.qty,
    c.normalPrice, c.normalTotal, c.soldTotal, c.discount, c.payment, c.note, "", "", c.ts, Date.now(), deviceHash, "", "",
    c.stand, c.size, c.cur, c.rate, c.local, c.order];
}

function writeRows_(s, rows) {
  var start = s.getLastRow() + 1, n = rows.length;
  // Tekstformat først, slik at innhold som begynner med = + - @ aldri tolkes som formel
  s.getRange(start, 1, n, 4).setNumberFormat("@");
  s.getRange(start, 10, n, 4).setNumberFormat("@");
  s.getRange(start, 16, n, 3).setNumberFormat("@");
  s.getRange(start, 19, n, 3).setNumberFormat("@");
  s.getRange(start, 24, n, 1).setNumberFormat("@");
  s.getRange(start, 1, n, SALES_HEADERS.length).setValues(rows);
}

function hash_(token) {
  token = String(token || "");
  if (!token) return "";
  return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, token));
}
function editWindowMin_() {
  var v = Number(PropertiesService.getScriptProperties().getProperty("EDIT_WINDOW_MINUTES"));
  return v > 0 ? v : DEFAULT_EDIT_MINUTES;
}

function readNames_(name, headers) {
  var s = sheet_(name, headers), n = s.getLastRow() - 1, out = [];
  if (n > 0) s.getRange(2, 1, n, 1).getValues().forEach(function (r) { if (r[0] !== "") out.push(String(r[0])); });
  return out;
}

function sizes_(v) {
  var seen = {}, out = [];
  String(v == null ? "" : v).split(/[,;\/]+/).forEach(function (x) {
    x = txt_(x, 30);
    if (x && !seen[x.toLowerCase()]) { seen[x.toLowerCase()] = true; out.push(x); }
  });
  return out.slice(0, 20);
}

function readStands_() {
  var s = standsSheet_(), n = s.getLastRow() - 1, out = [];
  if (n > 0) s.getRange(2, 1, n, 3).getValues().forEach(function (r) {
    if (r[0] === "") return;
    var cur = txt_(r[1], 8).toUpperCase() || "NOK", rate = num_(r[2]) > 0 ? num_(r[2]) : 1;
    out.push({ name: String(r[0]), cur: cur, rate: cur === "NOK" ? 1 : rate });
  });
  return out;
}

function readStock_() {
  var s = stockSheet_(), n = s.getLastRow() - 1, out = [];
  if (n > 0) s.getRange(2, 1, n, 4).getValues().forEach(function (r) {
    if (r[0] === "" || r[2] === "" || !isFinite(Number(r[2]))) return;
    out.push({ product: String(r[0]), size: String(r[1]), count: Math.round(Number(r[2])), ts: num_(r[3]) });
  });
  return out;
}

function readProducts_() {
  var s = productsSheet_(), n = s.getLastRow() - 1, out = [];
  if (n > 0) s.getRange(2, 1, n, 5).getValues().forEach(function (r) {
    if (r[0] !== "") out.push({ name: String(r[0]), normalPrice: Math.max(0, num_(r[1])), sizes: sizes_(r[3]),
      image: /^data:image\/jpeg;base64,/.test(String(r[4])) ? String(r[4]) : "" });
  });
  return out;
}

function readAll_() {
  var s = salesSheet_(), n = s.getLastRow() - 1, sales = [], deleted = [];
  if (n > 0) {
    s.getRange(2, 1, n, SALES_HEADERS.length).getValues().forEach(function (r) {
      if (r[0] === "") return;
      if (r[11] !== "") { deleted.push(String(r[0])); return; }
      sales.push({ id: String(r[0]), ts: num_(r[13]), seller: String(r[2]), product: String(r[3]), qty: num_(r[4]),
        normalPrice: num_(r[5]), normalTotal: num_(r[6]), soldTotal: num_(r[7]), discount: num_(r[8]),
        payment: String(r[9]), note: String(r[10]),
        stand: String(r[18]), size: String(r[19]), cur: String(r[20]) || "NOK", rate: num_(r[21]) > 0 ? num_(r[21]) : 1,
        local: r[22] === "" ? num_(r[7]) : num_(r[22]), order: String(r[23]) });
    });
  }
  return { ok: true, sales: sales, deletedIds: deleted, editWindowMin: editWindowMin_(),
    sellers: readNames_(SHEET_SELLERS, ["Navn"]).sort(function (a, b) { return a.localeCompare(b, "no"); }),
    products: readProducts_().sort(function (a, b) { return a.name.localeCompare(b.name, "no"); }),
    stands: readStands_(), stock: readStock_(),
    serverTime: Date.now() };
}

/* ---------- Åpne handlinger (ingen passord) ---------- */

function push_(b) {
  var s = salesSheet_(), n = s.getLastRow() - 1, known = {}, rows = [], added = 0, skipped = 0, dh = hash_(b.device);
  if (n > 0) s.getRange(2, 1, n, 1).getValues().forEach(function (r) { known[String(r[0])] = true; });
  (b.sales || []).slice(0, MAX_PUSH).forEach(function (x) {
    var c = cleanSale_(x);
    // Kjente ID-er hoppes over, også slettede, så et slettet salg aldri dukker opp igjen
    if (!c || known[c.id]) { skipped++; return; }
    known[c.id] = true; rows.push(rowOf_(c, dh)); added++;
  });
  if (rows.length) writeRows_(s, rows);
  addNames_(sellersSheet_(), b.sellers);
  addProducts_(b.products);
  return { ok: true, added: added, skipped: skipped };
}

function addNames_(s, list) {
  var have = {}, n = s.getLastRow() - 1, add = [];
  if (n > 0) s.getRange(2, 1, n, 1).getValues().forEach(function (r) { have[String(r[0]).toLowerCase()] = true; });
  (list || []).slice(0, 100).forEach(function (raw) {
    var name = txt_(raw, 100);
    if (name && !have[name.toLowerCase()]) { have[name.toLowerCase()] = true; add.push([name]); }
  });
  if (add.length) {
    var start = s.getLastRow() + 1;
    s.getRange(start, 1, add.length, 1).setNumberFormat("@");
    s.getRange(start, 1, add.length, 1).setValues(add);
  }
}

function addProducts_(list) {
  var s = productsSheet_(), n = s.getLastRow() - 1, rows = n > 0 ? s.getRange(2, 1, n, 2).getValues() : [];
  var index = {}, add = [];
  rows.forEach(function (r, i) { index[String(r[0]).toLowerCase()] = i; });
  (list || []).slice(0, 100).forEach(function (p) {
    if (!p) return;
    var name = txt_(p.name, 100), price = Math.max(0, num_(p.normalPrice));
    if (!name) return;
    var key = name.toLowerCase();
    if (index[key] === undefined) { index[key] = -1; add.push([name, price]); }
    else if (index[key] >= 0 && num_(rows[index[key]][1]) === 0 && price > 0) s.getRange(index[key] + 2, 2, 1, 1).setValues([[price]]);
  });
  if (add.length) {
    var start = s.getLastRow() + 1;
    s.getRange(start, 1, add.length, 1).setNumberFormat("@");
    s.getRange(start, 1, add.length, 2).setValues(add);
  }
}

/* ---------- Selgerens egne salg: endre og slette uten passord innenfor tidsgrensen ---------- */

// Finner raden og sjekker eierskap og tid. Tiden regnes fra det tidligste av telefonens tid og serverens mottakstid.
function ownRow_(s, id, device) {
  var n = s.getLastRow() - 1, dh = hash_(device);
  if (n < 1 || !id) return { error: "finnes_ikke" };
  var vals = s.getRange(2, 1, n, SALES_HEADERS.length).getValues();
  for (var i = 0; i < n; i++) {
    var r = vals[i];
    if (String(r[0]) !== String(id)) continue;
    if (r[11] !== "") return { error: "finnes_ikke" };
    if (!dh || String(r[15]) !== dh) return { error: "ikke_din" };
    var start = Math.min(num_(r[13]) || Infinity, num_(r[14]) || Infinity);
    if (!isFinite(start)) start = 0;
    var minutes = editWindowMin_();
    if (Date.now() - start > minutes * 60000) return { error: "utlopt", minutes: minutes };
    return { index: i + 2, row: r };
  }
  return { error: "finnes_ikke" };
}

function ownDelete_(b) {
  var s = salesSheet_(), f = ownRow_(s, b.id, b.device);
  if (!f.index) return { ok: false, error: f.error, minutes: f.minutes };
  s.getRange(f.index, 12, 1, 2).setValues([[stamp_(), "selger: " + txt_(b.by, 50)]]);
  return { ok: true, deleted: 1 };
}

function ownEdit_(b) {
  var s = salesSheet_(), f = ownRow_(s, b.id, b.device);
  if (!f.index) return { ok: false, error: f.error, minutes: f.minutes };
  var r = f.row;
  var c = cleanSale_({ id: r[0], ts: r[13], seller: r[2], product: b.product, qty: b.qty, normalPrice: b.normalPrice,
    soldTotal: b.soldTotal, payment: b.payment, note: b.note, stand: r[18], size: r[19], cur: r[20], rate: r[21], local: b.local, order: r[23] });
  if (!c) return { ok: false, error: "ugyldig" };
  var np = c.normalPrice;
  c.normalTotal = np > 0 ? c.qty * np : c.soldTotal;
  c.discount = c.normalTotal - c.soldTotal;
  var before = "Produkt=" + r[3] + ", Antall=" + r[4] + ", Normalpris=" + r[5] + ", Solgt=" + r[7] + ", Betaling=" + r[9] + ", Kommentar=" + r[10];
  var log = (String(r[17]) ? String(r[17]) + " | " : "") + stamp_() + " før: " + before;
  s.getRange(f.index, 4, 1, 8).setValues([[c.product, c.qty, c.normalPrice, c.normalTotal, c.soldTotal, c.discount, c.payment, c.note]]);
  s.getRange(f.index, 17, 1, 2).setValues([[stamp_(), log.slice(-1500)]]);
  s.getRange(f.index, 23, 1, 1).setValues([[c.local]]);
  return { ok: true, sale: { id: c.id, ts: c.ts, seller: c.seller, product: c.product, qty: c.qty, normalPrice: c.normalPrice,
    normalTotal: c.normalTotal, soldTotal: c.soldTotal, discount: c.discount, payment: c.payment, note: c.note,
    stand: c.stand, size: c.size, cur: c.cur, rate: c.rate, local: c.local, order: c.order } };
}

/* ---------- Handlinger som krever passord ---------- */

function stamp_() { return Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd HH:mm"); }

function deleteSales_(b) {
  var s = salesSheet_(), n = s.getLastRow() - 1, count = 0;
  var want = {};
  (b.ids || []).slice(0, 500).forEach(function (id) { want[String(id)] = true; });
  if (n > 0) {
    var ids = s.getRange(2, 1, n, 1).getValues(), del = s.getRange(2, 12, n, 1).getValues();
    for (var i = 0; i < n; i++) {
      if (want[String(ids[i][0])] && del[i][0] === "") {
        s.getRange(i + 2, 12, 1, 2).setValues([[stamp_(), txt_(b.by, 60)]]);
        count++;
      }
    }
  }
  return { ok: true, deleted: count };
}

function deleteAll_(b) {
  var s = salesSheet_(), n = s.getLastRow() - 1, count = 0;
  if (n > 0) {
    var ids = s.getRange(2, 1, n, 1).getValues(), del = s.getRange(2, 12, n, 1).getValues();
    for (var i = 0; i < n; i++) {
      if (ids[i][0] !== "" && del[i][0] === "") {
        s.getRange(i + 2, 12, 1, 2).setValues([[stamp_(), txt_(b.by, 60)]]);
        count++;
      }
    }
  }
  return { ok: true, deleted: count };
}

function setPrice_(b) {
  var s = productsSheet_(), n = s.getLastRow() - 1, name = txt_(b.name, 100).toLowerCase(), price = Math.max(0, num_(b.price));
  if (n > 0) {
    var names = s.getRange(2, 1, n, 1).getValues();
    for (var i = 0; i < n; i++) {
      if (String(names[i][0]).toLowerCase() === name) {
        if (b.price != null) s.getRange(i + 2, 2, 1, 1).setValues([[price]]);
        if (b.sizes !== undefined) {
          var list = sizes_(Array.isArray(b.sizes) ? b.sizes.join(",") : b.sizes).join(", ");
          s.getRange(i + 2, 4, 1, 1).setNumberFormat("@");
          s.getRange(i + 2, 4, 1, 1).setValues([[list]]);
        }
        return { ok: true };
      }
    }
  }
  return { ok: false, error: "finnes_ikke" };
}

function removeName_(sheetName, rawName) {
  var s = sheetName === SHEET_PRODUCTS ? productsSheet_() : sheetName === SHEET_STANDS ? standsSheet_() : sellersSheet_(), n = s.getLastRow() - 1, name = txt_(rawName, 100).toLowerCase();
  if (n > 0) {
    var names = s.getRange(2, 1, n, 1).getValues();
    for (var i = 0; i < n; i++) {
      if (String(names[i][0]).toLowerCase() === name) { s.deleteRow(i + 2); return { ok: true }; }
    }
  }
  return { ok: true };
}


/* ---------- Produktbilde ---------- */

function setImage_(b) {
  var s = productsSheet_(), n = s.getLastRow() - 1, name = txt_(b.name, 100).toLowerCase(), img = String(b.image || "");
  if (img && (img.indexOf("data:image/jpeg;base64,") !== 0 || img.length > 45000)) return { ok: false, error: "for_stort" };
  if (n > 0) {
    var names = s.getRange(2, 1, n, 1).getValues();
    for (var i = 0; i < n; i++) {
      if (String(names[i][0]).toLowerCase() === name) {
        s.getRange(i + 2, 5, 1, 1).setNumberFormat("@");
        s.getRange(i + 2, 5, 1, 1).setValues([[img]]);
        return { ok: true };
      }
    }
  }
  return { ok: false, error: "finnes_ikke" };
}

/* ---------- Stand (arrangement) og lager ---------- */

function setStand_(b) {
  var name = txt_(b.name, 60), cur = txt_(b.cur, 8).toUpperCase() || "NOK", rate = num_(b.rate);
  if (!name) return { ok: false, error: "ugyldig" };
  if (cur === "NOK") rate = 1;
  if (!(rate > 0)) return { ok: false, error: "ugyldig" };
  var s = standsSheet_(), n = s.getLastRow() - 1, names = n > 0 ? s.getRange(2, 1, n, 1).getValues() : [];
  for (var i = 0; i < n; i++) {
    if (String(names[i][0]).toLowerCase() === name.toLowerCase()) {
      s.getRange(i + 2, 2, 1, 2).setNumberFormat("@");
      s.getRange(i + 2, 2, 1, 1).setValues([[cur]]);
      s.getRange(i + 2, 3, 1, 1).setNumberFormat("0.0000");
      s.getRange(i + 2, 3, 1, 1).setValues([[rate]]);
      return { ok: true };
    }
  }
  var start = s.getLastRow() + 1;
  s.getRange(start, 1, 1, 2).setNumberFormat("@");
  s.getRange(start, 1, 1, 3).setValues([[name, cur, rate]]);
  return { ok: true };
}

// Varetelling: oppgitt antall gjelder fra tidspunktet for tellingen. Salg etter det trekkes fra i appen.
function setStock_(b) {
  var s = stockSheet_(), n = s.getLastRow() - 1, rows = n > 0 ? s.getRange(2, 1, n, 2).getValues() : [];
  var index = {}, now = Date.now(), done = 0;
  rows.forEach(function (r, i) { index[String(r[0]).toLowerCase() + "|" + String(r[1]).toLowerCase()] = i + 2; });
  (b.items || []).slice(0, 300).forEach(function (it) {
    if (!it) return;
    var p = txt_(it.product, 100), sz = txt_(it.size, 30), cnt = Math.round(num_(it.count));
    if (!p || it.count === "" || it.count == null || !isFinite(Number(it.count))) return;
    var key = p.toLowerCase() + "|" + sz.toLowerCase(), row = index[key];
    if (!row) { row = Math.max(s.getLastRow(), 1) + 1; index[key] = row; s.getRange(row, 1, 1, 2).setNumberFormat("@"); s.getRange(row, 1, 1, 2).setValues([[p, sz]]); }
    s.getRange(row, 3, 1, 3).setValues([[cnt, now, stamp_()]]);
    done++;
  });
  return { ok: true, updated: done };
}
