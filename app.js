/* Shopfloor Hausen-Berlin: UI, Microsoft sign-in and Excel (Graph) access. */
(function () {
  "use strict";

  var CFG = null, SCHEMA = null, msalApp = null, account = null;
  var D = { orders: [], shipments: [], problems: [], meetings: [], stock: {}, stockList: [], stockErr: "", loadedAt: null };
  var META = {}, ITEMS = {};
  var view = "meeting", filt = { orders: "offen", problems: "offen", q: "" };
  var M = { step: 0, start: null, stepStart: null };
  var loading = false, toastTimer = null;
  var $ = function (id) { return document.getElementById(id); };
  var S = window.SFL;

  /* ---------- small helpers ---------- */

  function esc(v) {
    return String(v === null || v === undefined ? "" : v).replace(/[&<>"']/g, function (c) {
      return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c];
    });
  }
  function today() { return S.isoOf(new Date()); }
  function userName() { return account ? (account.name || account.username || "") : ""; }
  function fd(iso) { return S.fmtDate(S.normDate(iso)); }
  function qty(n, unit) { return esc(String(Math.round(n * 1000) / 1000).replace(".", ",")) + (unit ? " " + esc(unit) : ""); }

  function toast(msg, err) {
    var t = $("toast");
    t.textContent = msg; t.className = err ? "err" : ""; t.style.display = "block";
    clearTimeout(toastTimer);
    toastTimer = setTimeout(function () { t.style.display = "none"; }, err ? 9000 : 3500);
  }
  function busy(on) { $("busy").style.display = on ? "block" : "none"; }
  function fail(e) { busy(false); console.error(e); toast(String(e && e.message ? e.message : e), true); }

  function getJSON(url) {
    return fetch(url, { cache: "no-store" }).then(function (r) {
      if (!r.ok) throw new Error(url + ": HTTP " + r.status);
      return r.json();
    });
  }

  /* ---------- sign-in (MSAL, redirect flow; works in Safari) ---------- */

  function redirectUri() {
    return CFG.redirect_uri || (location.origin + location.pathname.replace(/index\.html$/, ""));
  }

  function getToken() {
    var req = { scopes: CFG.scopes, account: account };
    return msalApp.acquireTokenSilent(req).then(function (r) { return r.accessToken; }, function (e) {
      if (e instanceof msal.InteractionRequiredAuthError || /interaction_required|login_required|consent_required/.test(e.errorCode || "")) {
        msalApp.acquireTokenRedirect(req);
        return new Promise(function () {});
      }
      throw e;
    });
  }

  /* ---------- Graph / Excel ---------- */

  function graph(method, path, body, attempt) {
    attempt = attempt || 0;
    return getToken().then(function (tok) {
      return fetch(CFG.graph_base.replace(/\/$/, "") + path, {
        method: method,
        headers: { "Authorization": "Bearer " + tok, "Content-Type": "application/json" },
        body: body ? JSON.stringify(body) : undefined
      });
    }).then(function (r) {
      if ((r.status === 429 || r.status === 503 || r.status === 504) && attempt < 2) {
        var wait = (parseInt(r.headers.get("Retry-After"), 10) || 2) * 1000;
        return new Promise(function (res) { setTimeout(res, wait); }).then(function () { return graph(method, path, body, attempt + 1); });
      }
      return r.text().then(function (t) {
        if (!r.ok) throw new Error("Graph " + r.status + ": " + t.slice(0, 300));
        return t ? JSON.parse(t) : {};
      });
    });
  }

  function shareId(link) {
    var b = btoa(unescape(encodeURIComponent(link.trim())));
    return "u!" + b.replace(/=+$/, "").replace(/\//g, "_").replace(/\+/g, "-");
  }

  function workbook(link) {
    if (ITEMS[link]) return Promise.resolve(ITEMS[link]);
    return graph("GET", "/shares/" + shareId(link) + "/driveItem").then(function (it) {
      ITEMS[link] = "/drives/" + it.parentReference.driveId + "/items/" + it.id + "/workbook";
      return ITEMS[link];
    });
  }

  function colNum(letters) {
    var n = 0;
    for (var i = 0; i < letters.length; i++) n = n * 26 + (letters.charCodeAt(i) - 64);
    return n - 1;
  }
  function colLetter(idx) {
    var s = ""; idx += 1;
    while (idx) { var r = (idx - 1) % 26; s = String.fromCharCode(65 + r) + s; idx = Math.floor((idx - 1) / 26); }
    return s;
  }
  function parseAddr(a) {
    var m = /^(?:'?(.*?)'?)!\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$/.exec(a);
    if (!m) throw new Error("Unerwartete Bereichsadresse: " + a);
    return { sheet: m[1], c1: colNum(m[2]), r1: parseInt(m[3], 10), c2: colNum(m[4] || m[2]) };
  }

  function normCell(type, v) {
    if (type === "date") return S.normDate(v);
    if (type === "num") return S.hasNum(v) ? S.num(v) : S.str(v);
    return S.str(v);
  }

  /* Read one Excel table; columns are matched by header name (case-insensitive). */
  function readTable(key) {
    var t = SCHEMA[key];
    return workbook(CFG.shopfloor_share_link).then(function (wb) {
      return graph("GET", wb + "/tables('" + encodeURIComponent(t.table) + "')/range").then(function (rg) {
        var ad = parseAddr(rg.address), vals = rg.values || [];
        var header = (vals[0] || []).map(function (h) { return S.str(h).toLowerCase(); });
        var colKey = header.map(function () { return null; }), colType = {};
        t.columns.forEach(function (c) {
          var i = header.indexOf(c[1].toLowerCase());
          if (i < 0 && c[3] === "optional") return;
          if (i < 0) throw new Error("Spalte '" + c[1] + "' fehlt in Tabelle " + t.table);
          colKey[i] = c[0]; colType[c[0]] = c[2];
        });
        META[key] = { sheet: ad.sheet, c1: ad.c1, c2: ad.c2, r1: ad.r1, colKey: colKey, colType: colType };
        var rows = [];
        vals.slice(1).forEach(function (arr, i) {
          if (!arr.some(function (v) { return S.str(v) !== ""; })) return;
          var o = { _row: ad.r1 + 1 + i, _raw: arr };
          colKey.forEach(function (k, ci) { if (k) o[k] = normCell(colType[k], arr[ci]); });
          rows.push(o);
        });
        return rows;
      });
    });
  }

  /* Value for a cell write; text that Excel would treat as a formula is kept as text. */
  function cellOut(type, v) {
    if (v === null || v === undefined) return "";
    if (type === "num") return S.hasNum(v) ? S.num(v) : "";
    if (type === "jn") return v === true ? "J" : v === false ? "N" : S.str(v);
    var s = S.str(v);
    /* "'" prefix keeps formulas, leading zeros (e.g. OrgaMax "00123") and fraction/date-like text (LB-Nr "1234/26") as text */
    if (type === "text" && (/^[=+@]/.test(s) || /^0\d+$/.test(s) || /^\d+\/\d+$/.test(s))) s = "'" + s;
    return s;
  }

  function rowValues(key, obj, raw) {
    var m = META[key];
    return m.colKey.map(function (k, ci) {
      if (k && Object.prototype.hasOwnProperty.call(obj, k)) return cellOut(m.colType[k], obj[k]);
      return raw ? raw[ci] : "";
    });
  }

  function addRows(key, objs) {
    return workbook(CFG.shopfloor_share_link).then(function (wb) {
      return graph("POST", wb + "/tables('" + encodeURIComponent(SCHEMA[key].table) + "')/rows/add",
        { index: null, values: objs.map(function (o) { return rowValues(key, o, null); }) });
    });
  }

  /* Re-read the table, find exactly one matching row, then write only the changed fields. */
  function updateRow(key, match, patch) {
    return readTable(key).then(function (rows) {
      var hit = rows.filter(match);
      if (hit.length !== 1) throw new Error("Datensatz nicht eindeutig gefunden (" + hit.length + ") – bitte Aktualisieren.");
      var r = hit[0], m = META[key];
      var addr = colLetter(m.c1) + r._row + ":" + colLetter(m.c2) + r._row;
      return workbook(CFG.shopfloor_share_link).then(function (wb) {
        return graph("PATCH", wb + "/worksheets('" + encodeURIComponent(m.sheet) + "')/range(address='" + addr + "')",
          { values: [rowValues(key, patch, r._raw)] });
      });
    });
  }

  /* Hausen stock, read-only, same sheet and columns as the Pi kiosk app. */
  function loadStock() {
    var st = CFG.stock || {};
    if (!st.share_link) return Promise.resolve({ map: {}, list: [], err: "" });
    return workbook(st.share_link).then(function (wb) {
      return graph("GET", wb + "/worksheets('" + encodeURIComponent(st.sheet_name) + "')/usedRange(valuesOnly=true)?$select=address,values");
    }).then(function (ur) {
      var ad = parseAddr(ur.address), vals = ur.values || [];
      var hi = (st.header_row || 1) - ad.r1;
      var header = (vals[hi] || []).map(function (h) { return S.str(h).toLowerCase(); });
      function find(name) { return name ? header.indexOf(String(name).toLowerCase()) : -1; }
      var ci = { id: find(st.id_column), name: find(st.name_column), qty: find(st.qty_column),
                 min: find(st.min_column), unit: find(st.unit_column) };
      if (ci.id < 0 || ci.qty < 0) throw new Error("Lagerliste: Spalte '" + st.id_column + "' oder '" + st.qty_column + "' nicht gefunden");
      var map = {}, list = [];
      vals.slice(hi + 1).forEach(function (r) {
        var id = S.str(r[ci.id]);
        if (!id) return;
        var it = { snr: id, name: ci.name >= 0 ? S.str(r[ci.name]) : "", qty: S.num(r[ci.qty]),
                   min: ci.min >= 0 && S.hasNum(r[ci.min]) ? S.num(r[ci.min]) : null,
                   unit: ci.unit >= 0 ? S.str(r[ci.unit]) : "" };
        map[id] = it; list.push(it);
      });
      return { map: map, list: list, err: "" };
    }).catch(function (e) {
      return { map: {}, list: [], err: String(e.message || e) };
    });
  }

  function loadAll(quiet) {
    if (loading) return Promise.resolve();
    loading = true; busy(true);
    return Promise.all([readTable("orders"), readTable("shipments"), readTable("problems"), readTable("meetings"), loadStock()])
      .then(function (r) {
        D.orders = r[0]; D.shipments = r[1]; D.problems = r[2]; D.meetings = r[3];
        D.stock = r[4].map; D.stockList = r[4].list; D.stockErr = r[4].err;
        D.loadedAt = new Date();
        loading = false; busy(false);
        $("info").textContent = userName() + " · Stand " + D.loadedAt.toLocaleTimeString("de-DE", { hour: "2-digit", minute: "2-digit" });
        if (!quiet || !document.querySelector("#view form")) render();
      }, function (e) { loading = false; fail(e); });
  }

  /* ---------- rendering helpers ---------- */

  function lines() { return S.deriveLines(D.orders, D.shipments, today(), CFG.due_soon_days); }
  function lineMap(ls) { var m = {}; ls.forEach(function (l) { m[l.key] = l; }); return m; }

  var FLAG_CLS = { "Überfällig": "bad", "Bald fällig": "bad", "QS n.i.O.": "bad", "Mengenabweichung": "warn", "QS offen": "warn",
                   "Verspätet": "warn", "Express": "info" };
  function chip(text, cls) { return '<span class="chip ' + (cls || "") + '">' + esc(text) + "</span>"; }
  function lineChips(l) {
    var c = l.done ? "good" : "";
    return chip(l.status, c) + l.flags.map(function (f) { return chip(f, FLAG_CLS[f]); }).join("");
  }
  /* Card background = schedule (overdue / due soon / partly delivered), left border = quantity/QS issues. */
  function lineCls(l) {
    var bg = S.hasFlag(l, "Überfällig") ? " late" : S.hasFlag(l, "Bald fällig") ? " soon" : (!l.done && l.received > 0) ? " part" : "";
    if (S.hasFlag(l, "Überfällig") || S.hasFlag(l, "QS n.i.O.")) return "bad" + bg;
    if (S.hasFlag(l, "Mengenabweichung") || S.hasFlag(l, "QS offen") || S.hasFlag(l, "Verspätet")) return "warn" + bg;
    return (l.done ? "good" : "") + bg;
  }
  function stockHtml(snr) {
    if (!snr) return '<span class="muted">–</span>';
    if (D.stockErr) return '<span class="muted">n/a</span>';
    var s = D.stock[snr];
    if (!s) return '<span class="muted">nicht in Lagerliste</span>';
    var cls = s.qty <= 0 ? "zero" : (s.min !== null && s.qty < s.min ? "low" : "");
    return '<span class="stock ' + cls + '">' + qty(s.qty, s.unit) + (s.min !== null ? " (Min " + qty(s.min) + ")" : "") + "</span>";
  }
  function g(label, val) { return "<div><b>" + esc(label) + "</b>" + val + "</div>"; }

  function lineCard(l, acts) {
    var u = l.einheit;
    return '<div class="item ' + lineCls(l) + '"><div class="top"><span class="title">' + esc(l.bestellnr) + " / " + esc(l.pos) +
      " – " + esc(l.artikel || "?") + "</span>" + lineChips(l) + '</div><div class="grid">' +
      g("S-Nr", esc(l.snr || "–")) + (l.lbnr ? g("LB-Nr", esc(l.lbnr)) : "") + g("Bestellt", qty(l.ordered, u)) + g("Erhalten (i.O.)", qty(l.received, u) + " (" + qty(l.okQty) + ")") +
      g("Unterwegs", qty(l.transit, u)) + g("Offen", qty(l.open, u)) +
      g("Termin", esc(S.fmtDate(l.due) || "–") + (S.normDate(l.order.bestaetigt) ? " ✓" : "")) +
      g("Lager Hausen", stockHtml(l.snr)) +
      (l.order.bemerkung ? g("Bemerkung", esc(l.order.bemerkung)) : "") +
      '</div><div class="acts">' + (acts || "") + "</div></div>";
  }

  function shipArtikel(s, lm) {
    var l = s.bestellnr ? lm[S.lineKey(s.bestellnr, s.pos)] : null;
    return l ? l.artikel : s.artikel;
  }
  function shipCls(s) {
    var r = s.qs_ergebnis;
    if (!s.bestellnr && s.richtung === S.DIR_BH) return "bad";
    if (S.isReceived(s)) {
      if (r === "n.i.O." || !r) return "bad";
      if (r === "teilweise" || (S.hasNum(s.menge_versendet) && S.num(s.menge_versendet) !== S.num(s.menge_erhalten))) return "warn";
      return "good";
    }
    return "";
  }
  function shipCard(s, lm, acts) {
    var rec = S.isReceived(s), r = s.qs_ergebnis, chips = "";
    chips += chip(s.richtung || "–", "info");
    if (S.yes(s.express)) chips += chip("Express", "info");
    if (!s.bestellnr) chips += chip("ohne Bestellung", s.richtung === S.DIR_BH ? "bad" : "warn");
    if (!rec) {
      var d = S.daysBetween(S.normDate(s.versanddatum), today());
      chips += chip("unterwegs" + (s.versanddatum ? " seit " + d + " T." : ""), d > 3 ? "warn" : "");
    } else {
      chips += r ? chip("QS " + r, r === "i.O." ? "good" : r === "n.i.O." ? "bad" : "warn") : chip("QS offen", "bad");
      if (S.hasNum(s.menge_versendet) && S.num(s.menge_versendet) !== S.num(s.menge_erhalten)) chips += chip("Mengenabweichung", "warn");
    }
    var l = s.bestellnr ? lm[S.lineKey(s.bestellnr, s.pos)] : null;
    return '<div class="item ' + shipCls(s) + '"><div class="top"><span class="title">' + esc(s.lieferid) + " – " +
      esc(shipArtikel(s, lm) || "?") + "</span>" + chips + '</div><div class="grid">' +
      g("Bestellung", s.bestellnr ? esc(s.bestellnr + " / " + s.pos) + (l ? " (bestellt " + qty(l.ordered) + ")" : "") : "–") +
      g("Versendet", s.versanddatum || S.hasNum(s.menge_versendet) ? esc(fd(s.versanddatum) || "?") + ": " + (S.hasNum(s.menge_versendet) ? qty(s.menge_versendet) : "?") : "keine Versandmeldung") +
      g("Versendet von", esc(s.versendet_von || "–")) +
      (s.tracking ? g("Paket/Tracking", esc(s.tracking)) : "") +
      (rec ? g("Eingang", esc(fd(s.eingangsdatum)) + ": " + qty(s.menge_erhalten)) +
             g("Menge i.O.", S.hasNum(s.menge_io) ? qty(s.menge_io) : "–") +
             g("QS geprüft von", esc(s.qs_von || "–")) : "") +
      (s.qs_bemerkung ? g("QS-Bemerkung", esc(s.qs_bemerkung)) : "") +
      (s.bemerkung ? g("Bemerkung", esc(s.bemerkung)) : "") +
      '</div><div class="acts">' + (acts || "") + "</div></div>";
  }

  var URG = { hoch: 0, mittel: 1, niedrig: 2 };
  function probCard(p) {
    var late = p.status !== "erledigt" && p.termin && p.termin < today();
    var cls = p.status === "erledigt" ? "good" : (p.dringlichkeit === "hoch" || late ? "bad" : "warn");
    return '<div class="item ' + cls + '"><div class="top"><span class="title">#' + esc(p.nr) + " " + esc(p.problem) + "</span>" +
      chip(p.status || "offen", p.status === "erledigt" ? "good" : "") + (p.kategorie ? chip(p.kategorie, "info") : "") +
      (p.dringlichkeit ? chip(p.dringlichkeit, p.dringlichkeit === "hoch" ? "bad" : "") : "") + (late ? chip("Termin überschritten", "bad") : "") +
      '</div><div class="grid">' + g("Bezug", esc(p.bezug || "–")) + g("Verantwortlich", esc(p.verantwortlich || "–")) +
      g("Standort", esc(p.standort || "–")) + g("Erfasst", esc(fd(p.datum))) + g("Termin", esc(fd(p.termin) || "–")) +
      (p.massnahme ? g("Maßnahme", esc(p.massnahme)) : "") +
      '</div><div class="acts">' +
      '<button class="btn sm" data-act="editProblem" data-nr="' + esc(p.nr) + '">Bearbeiten</button>' +
      (p.status !== "in Arbeit" && p.status !== "erledigt" ? '<button class="btn sm" data-act="probStatus" data-nr="' + esc(p.nr) + '" data-s="in Arbeit">→ in Arbeit</button>' : "") +
      (p.status !== "erledigt" ? '<button class="btn sm ok" data-act="probStatus" data-nr="' + esc(p.nr) + '" data-s="erledigt">✓ erledigt</button>' : "") +
      "</div></div>";
  }
  function openProblems() {
    return D.problems.filter(function (p) { return p.problem && p.status !== "erledigt"; }).sort(function (a, b) {
      var u = (URG[a.dringlichkeit] === undefined ? 1 : URG[a.dringlichkeit]) - (URG[b.dringlichkeit] === undefined ? 1 : URG[b.dringlichkeit]);
      if (u) return u;
      return (a.termin || "9999") < (b.termin || "9999") ? -1 : 1;
    });
  }

  function list(items, emptyText) {
    return items.length ? '<div class="list">' + items.join("") + "</div>" : '<div class="empty">' + esc(emptyText) + "</div>";
  }
  function probBtn(bezug, kat) {
    return '<button class="btn sm" data-act="newProblem" data-bezug="' + esc(bezug) + '" data-kat="' + esc(kat || "") + '">+ Problem</button>';
  }
  function options(arr, sel, blank) {
    return (blank !== undefined ? '<option value="">' + esc(blank) + "</option>" : "") + arr.map(function (v) {
      return '<option value="' + esc(v) + '"' + (v === sel ? " selected" : "") + ">" + esc(v) + "</option>";
    }).join("");
  }
  function seg(name, arr, sel) {
    return '<input type="hidden" name="' + esc(name) + '" value="' + esc(sel || "") + '"><div class="seg">' + arr.map(function (v) {
      return '<button type="button" data-act="seg" data-name="' + esc(name) + '" data-val="' + esc(v) + '"' + (v === sel ? ' class="on"' : "") + ">" + esc(v) + "</button>";
    }).join("") + "</div>";
  }
  function field(label, inner) { return '<div class="f"><label>' + esc(label) + "</label>" + inner + "</div>"; }
  function openLineOptions(ls, sel) {
    return S.sortLines(ls.filter(function (l) { return !l.done; })).map(function (l) {
      return '<option value="' + esc(l.key) + '"' + (l.key === sel ? " selected" : "") + ">" + esc(l.bestellnr + " / " + l.pos + " – " + l.artikel +
        " – offen " + (l.open - l.transit) + " " + l.einheit + (l.due ? " – Termin " + S.fmtDate(l.due) : "")) + "</option>";
    }).join("");
  }

  /* ---------- views ---------- */

  var STEPS = ["Kennzahlen", "Offene Bestellungen", "Unterwegs", "Eingänge & QS", "Probleme", "Abschluss"];

  function viewMeeting() {
    var ls = lines(), lm = lineMap(ls), per = S.meetingPeriod(D.meetings, today());
    if (!M.start) { M.start = Date.now(); M.stepStart = Date.now(); }
    var h = '<div class="bar"><h2 class="grow" style="margin:0">Shopfloor-Meeting – ' + esc(S.fmtDate(today())) + "</h2>" +
      '<span class="muted">Zeitraum ab ' + esc(S.fmtDate(per.from)) + (per.last ? " (letztes Meeting)" : " (7 Tage)") + "</span>" +
      '<span id="mtimer"></span><button class="btn sm" data-act="mreset">Neu starten</button></div>';
    h += '<div class="steps">' + STEPS.map(function (t, i) {
      return '<div title="' + esc(t) + '" data-act="mstep" data-i="' + i + '" class="' + (i === M.step ? "on" : i < M.step ? "done" : "") + '"></div>';
    }).join("") + "</div>";
    h += '<div class="bar"><h3 class="grow" style="margin:0">' + (M.step + 1) + "/" + STEPS.length + " " + esc(STEPS[M.step]) + "</h3>" +
      '<button class="btn" data-act="mstep" data-i="' + (M.step - 1) + '"' + (M.step === 0 ? " disabled" : "") + ">← Zurück</button>" +
      (M.step < STEPS.length - 1 ? '<button class="btn pri" data-act="mstep" data-i="' + (M.step + 1) + '">Weiter →</button>' : "") + "</div>";

    if (M.step === 0) {
      var ks = S.kpis(ls, D.shipments, D.problems, per.from, per.to, today(), CFG.kpi_targets);
      h += '<div class="kpis">' + ks.map(function (k) {
        var tgt = k.target === undefined ? "" : "Ziel " + (k.dir === "min" ? "≥ " : "≤ ") + (k.pct ? Math.round(k.target * 100) + " %" : k.target);
        return '<div class="kpi ' + (k.ok === null ? "" : k.ok ? "good" : "bad") + '"><div class="l">' + esc(k.label) + '</div><div class="v">' +
          esc(S.fmtKpi(k)) + '</div><div class="h">' + esc(tgt) + "<br>" + esc(k.hint) + "</div></div>";
      }).join("") + "</div>";
      h += '<p class="muted">Liefertreue, Mengentreue, QS-Quote und Bestellbezug beziehen sich auf Eingänge Berlin → Hausen im Zeitraum.</p>';
    } else if (M.step === 1) {
      var open = S.sortLines(ls.filter(function (l) { return !l.done; }));
      h += '<p class="muted">' + open.length + " offene Positionen, davon " +
        open.filter(function (l) { return S.hasFlag(l, "Überfällig"); }).length + " überfällig.</p>";
      h += list(open.map(function (l) {
        return lineCard(l, '<button class="btn sm" data-act="editOrder" data-k="' + esc(l.key) + '">Termin/Bearbeiten</button>' +
          probBtn("Bestellung " + l.bestellnr + "/" + l.pos, "Termin"));
      }), "Keine offenen Bestellpositionen.");
    } else if (M.step === 2) {
      var tr = D.shipments.filter(function (s) { return !S.isReceived(s); }).sort(function (a, b) {
        return (a.versanddatum || "") < (b.versanddatum || "") ? -1 : 1;
      });
      h += list(tr.map(function (s) {
        return shipCard(s, lm, '<button class="btn sm pri" data-act="receive" data-id="' + esc(s.lieferid) + '">Eingang erfassen</button>' +
          probBtn("Lieferung " + s.lieferid, "Termin"));
      }), "Nichts unterwegs.");
    } else if (M.step === 3) {
      var rc = D.shipments.filter(function (s) { var d = S.normDate(s.eingangsdatum); return d && d >= per.from; })
        .sort(function (a, b) { return a.eingangsdatum < b.eingangsdatum ? 1 : -1; });
      h += '<p class="muted">Eingänge seit ' + esc(S.fmtDate(per.from)) + ": " + rc.length + ". Rot = ohne Bestellbezug, QS offen oder n.i.O.</p>";
      h += list(rc.map(function (s) {
        return shipCard(s, lm, (!s.qs_ergebnis ? '<button class="btn sm pri" data-act="receive" data-id="' + esc(s.lieferid) + '">QS erfassen</button>' : "") +
          probBtn("Lieferung " + s.lieferid, s.qs_ergebnis && s.qs_ergebnis !== "i.O." ? "Qualität" : "Menge"));
      }), "Keine Eingänge im Zeitraum.");
    } else if (M.step === 4) {
      h += '<div class="bar"><span class="grow muted">Offene Probleme, nach Dringlichkeit</span>' + probBtn("", "") + "</div>";
      h += list(openProblems().map(probCard), "Keine offenen Probleme.");
    } else {
      var ks2 = S.kpis(ls, D.shipments, D.problems, per.from, per.to, today(), CFG.kpi_targets);
      h += '<form data-form="saveMeeting">' +
        field("Teilnehmer", '<input name="teilnehmer" required value="' + esc(userName()) + '">') +
        field("Notizen / Vereinbarungen", '<textarea name="notizen"></textarea>') +
        '<p class="muted">Gespeichert werden Datum, Dauer und die Kennzahlen: ' + ks2.map(function (k) { return esc(k.label + " " + S.fmtKpi(k)); }).join(" · ") + "</p>" +
        '<button class="btn ok" type="submit">Meeting speichern</button></form>';
    }
    return h;
  }

  function orderList(ls) {
    var q = filt.q.toLowerCase();
    var sel = ls.filter(function (l) {
      if (filt.orders === "offen" && l.done) return false;
      if (filt.orders === "überfällig" && !S.hasFlag(l, "Überfällig")) return false;
      if (q && (l.bestellnr + " " + l.artikel + " " + l.snr + " " + l.lbnr).toLowerCase().indexOf(q) < 0) return false;
      return true;
    });
    sel = filt.orders === "alle" ? sel.sort(function (a, b) {
      return a.bestellnr === b.bestellnr ? a.pos - b.pos : (a.bestellnr < b.bestellnr ? 1 : -1);
    }) : S.sortLines(sel);
    return list(sel.map(function (l) {
      return lineCard(l, '<button class="btn sm" data-act="editOrder" data-k="' + esc(l.key) + '">Bearbeiten</button>' +
        probBtn("Bestellung " + l.bestellnr + "/" + l.pos, ""));
    }), "Keine Positionen für diesen Filter.");
  }

  function viewOrders() {
    var ls = lines();
    return '<div class="bar"><h2 class="grow" style="margin:0">Bestellungen Hausen → Berlin</h2>' +
      '<button class="btn pri" data-act="newOrder">+ Neue Bestellung</button></div>' +
      '<div class="bar"><div class="seg" style="flex:2">' + ["offen", "überfällig", "alle"].map(function (f) {
        return '<button type="button" data-act="ofilter" data-f="' + f + '"' + (filt.orders === f ? ' class="on"' : "") + ">" + esc(f[0].toUpperCase() + f.slice(1)) + "</button>";
      }).join("") + '</div><form style="flex:3;min-width:200px" onsubmit="return false"><input id="oq" placeholder="Suchen: BestellNr, Artikel, S-Nr, LB-Nr" value="' + esc(filt.q) + '"></form></div>' +
      '<div id="olist">' + orderList(ls) + "</div>";
  }

  function viewShip() {
    var ls = lines(), lm = lineMap(ls);
    var recent = D.shipments.slice().sort(function (a, b) { return (a.lieferid < b.lieferid) ? 1 : -1; }).slice(0, 15);
    return "<h2>Versand melden</h2>" +
      '<p class="muted">Jedes Paket hier melden – auch Express. Die LieferID danach auf Lieferschein und Paket schreiben.</p>' +
      '<form data-form="ship" class="item">' +
      field("Richtung", seg("richtung", SCHEMA.lists.richtung, S.DIR_BH)) +
      '<div data-dir="bh">' + field("Bestellposition", '<select name="line"><option value="">– bitte wählen –</option>' + openLineOptions(ls) +
        '<option value="-">Ohne Bestellung (Grund angeben)</option></select>') + "</div>" +
      '<div data-dir="free" style="display:none">' + field("Artikel", '<input name="artikel" list="stocklist">') +
        field("Grund / Bezug (Pflicht ohne Bestellung)", '<input name="bemerkung">') + "</div>" +
      '<div class="row2">' + field("Menge", '<input name="menge" type="number" min="0" step="any" inputmode="decimal" required>') +
        field("Versanddatum", '<input name="versanddatum" type="date" required value="' + today() + '">') + "</div>" +
      '<div class="row2">' + field("Paket/Tracking", '<input name="tracking">') +
        '<div class="f"><label>&nbsp;</label><label class="chk"><input type="checkbox" name="express"> Express</label></div></div>' +
      '<button class="btn ok" type="submit">Versand melden</button></form>' + stockDatalist() +
      "<h2>Letzte Meldungen</h2>" + list(recent.map(function (s) {
        return shipCard(s, lm, !S.isReceived(s) ? '<button class="btn sm" data-act="receive" data-id="' + esc(s.lieferid) + '">Eingang erfassen</button>' : "");
      }), "Noch keine Lieferungen.");
  }

  function viewIn() {
    var ls = lines(), lm = lineMap(ls), from = S.addDays(today(), -14);
    var tr = D.shipments.filter(function (s) { return !S.isReceived(s); }).sort(function (a, b) { return (a.versanddatum || "") < (b.versanddatum || "") ? -1 : 1; });
    var rc = D.shipments.filter(function (s) { var d = S.normDate(s.eingangsdatum); return d && d >= from; })
      .sort(function (a, b) { return a.eingangsdatum < b.eingangsdatum ? 1 : -1; });
    return '<div class="bar"><h2 class="grow" style="margin:0">Wareneingang & QS</h2>' +
      '<button class="btn pri" data-act="receiveNew">+ Paket ohne Versandmeldung</button></div>' +
      "<h3>Erwartet (unterwegs)</h3>" + list(tr.map(function (s) {
        return shipCard(s, lm, '<button class="btn sm pri" data-act="receive" data-id="' + esc(s.lieferid) + '">Eingang erfassen</button>');
      }), "Keine Pakete unterwegs.") +
      "<h3>Eingänge der letzten 14 Tage</h3>" + list(rc.map(function (s) {
        return shipCard(s, lm, '<button class="btn sm" data-act="receive" data-id="' + esc(s.lieferid) + '">' + (s.qs_ergebnis ? "Ändern" : "QS erfassen") + "</button>" +
          probBtn("Lieferung " + s.lieferid, "Qualität"));
      }), "Keine Eingänge.");
  }

  function viewProblems() {
    var ps = filt.problems === "offen" ? openProblems() :
      D.problems.filter(function (p) { return p.problem; }).sort(function (a, b) { return S.num(b.nr) - S.num(a.nr); });
    return '<div class="bar"><h2 class="grow" style="margin:0">Problemkarten</h2>' + probBtn("", "") + "</div>" +
      '<div class="bar"><div class="seg">' + ["offen", "alle"].map(function (f) {
        return '<button type="button" data-act="pfilter" data-f="' + f + '"' + (filt.problems === f ? ' class="on"' : "") + ">" + esc(f[0].toUpperCase() + f.slice(1)) + "</button>";
      }).join("") + "</div></div>" + list(ps.map(probCard), "Keine Problemkarten.");
  }

  function stockDatalist() {
    return '<datalist id="stocklist">' + D.stockList.map(function (s) {
      return '<option value="' + esc(s.name) + '">' + esc(s.snr) + "</option>";
    }).join("") + '</datalist><datalist id="snrlist">' + D.stockList.map(function (s) {
      return '<option value="' + esc(s.snr) + '">' + esc(s.name) + "</option>";
    }).join("") + "</datalist>";
  }

  function render() {
    var v = $("view"), h;
    Array.prototype.forEach.call(document.querySelectorAll("#nav a"), function (a) { a.className = a.getAttribute("data-v") === view ? "on" : ""; });
    if (view === "bestellungen") h = viewOrders();
    else if (view === "versand") h = viewShip();
    else if (view === "eingang") h = viewIn();
    else if (view === "probleme") h = viewProblems();
    else h = viewMeeting();
    if (D.stockErr) h = '<div class="item warn" style="margin-bottom:8px">Lagerliste Hausen nicht lesbar: ' + esc(D.stockErr) + "</div>" + h;
    v.innerHTML = h;
    tick();
  }

  function tick() {
    var el = $("mtimer");
    if (!el || !M.start) return;
    function mmss(ms) { var s = Math.floor(ms / 1000); return Math.floor(s / 60) + ":" + ("0" + (s % 60)).slice(-2); }
    var stepMs = Date.now() - M.stepStart;
    el.textContent = "Gesamt " + mmss(Date.now() - M.start) + " · Schritt " + mmss(stepMs);
    el.className = stepMs > (CFG.meeting_step_minutes || 2) * 60000 ? "over" : "";
  }

  /* ---------- modals and forms ---------- */

  function openModal(html) { $("modal").innerHTML = html; $("overlay").style.display = "flex"; }
  function closeModal() { $("overlay").style.display = "none"; $("modal").innerHTML = ""; }
  function modalShell(title, form, inner, submitText) {
    return "<h2>" + esc(title) + '</h2><form data-form="' + form + '">' + inner +
      '<div class="bar"><button class="btn ok" type="submit">' + esc(submitText || "Speichern") + '</button>' +
      '<button class="btn" type="button" data-act="close">Abbrechen</button></div></form>';
  }
  function formData(f) {
    var o = {};
    Array.prototype.forEach.call(f.elements, function (el) {
      if (!el.name) return;
      if (el.type === "checkbox") o[el.name] = el.checked; else o[el.name] = String(el.value || "").trim();
    });
    return o;
  }
  function orderLineRow() {
    return '<div class="ol"><input name="ol_snr" list="snrlist" placeholder="S-Nr" data-snr="1">' +
      '<input name="ol_artikel" placeholder="Artikel" list="stocklist"><input name="ol_menge" type="number" min="0" step="any" inputmode="decimal" placeholder="Menge">' +
      '<button class="btn sm" type="button" data-act="rmLine">✕</button></div>';
  }

  function lbField(v) {
    if (!META.orders || !META.orders.colType.lbnr) return "";
    return field("LB-Nr (Laufband, z. B. 1234/26, 1235/26)", '<input name="lbnr" value="' + esc(v) + '" placeholder="1234/26">');
  }
  function lbValue(d) {
    if (d.lbnr === undefined) return undefined;
    var p = S.parseLbNr(d.lbnr);
    if (!p.ok) throw new Error("LB-Nr ungültig: " + p.bad.join(", ") + " – Format 1234/26, mehrere mit Komma trennen.");
    return p.value;
  }

  function modalNewOrder() {
    openModal(modalShell("Neue Bestellung (aus OrgaMax)", "newOrder",
      '<div class="row2">' + field("BestellNr (OrgaMax)", '<input name="bestellnr" required>') +
      field("Wunschtermin", '<input name="wunschtermin" type="date" required>') + "</div>" +
      lbField("") +
      field("Positionen (S-Nr aus Lagerliste Hausen)", '<div class="olines" id="olines">' + orderLineRow() + "</div>" +
        '<button class="btn sm" type="button" data-act="addLine">+ Position</button>') +
      '<label class="chk f"><input type="checkbox" name="express"> Express</label>' +
      field("Bemerkung", '<input name="bemerkung">') + stockDatalist()));
  }

  function modalEditOrder(key) {
    var l = lines().filter(function (x) { return x.key === key; })[0];
    if (!l) return toast("Position nicht gefunden – bitte Aktualisieren.", true);
    var o = l.order;
    openModal(modalShell("Bestellung " + l.bestellnr + " / " + l.pos, "editOrder",
      '<input type="hidden" name="key" value="' + esc(key) + '"><p class="muted">' + esc(l.artikel) + " · S-Nr " + esc(l.snr || "–") + "</p>" +
      '<div class="row2">' + field("Menge", '<input name="menge" type="number" min="0" step="any" value="' + esc(o.menge) + '">') +
      field("Wunschtermin", '<input name="wunschtermin" type="date" value="' + esc(o.wunschtermin) + '">') + "</div>" +
      field("Bestätigter Termin (Berlin)", '<input name="bestaetigt" type="date" value="' + esc(o.bestaetigt) + '">') +
      lbField(o.lbnr || "") +
      '<label class="chk f"><input type="checkbox" name="express"' + (S.yes(o.express) ? " checked" : "") + "> Express</label>" +
      '<label class="chk f"><input type="checkbox" name="abgeschlossen"' + (S.yes(o.abgeschlossen) ? " checked" : "") + "> Abgeschlossen (Rest wird nicht mehr geliefert)</label>" +
      field("Bemerkung", '<input name="bemerkung" value="' + esc(o.bemerkung) + '">')));
  }

  function findShip(id) { return D.shipments.filter(function (s) { return s.lieferid === id; })[0]; }

  function modalReceive(id) {
    var s = findShip(id);
    if (!s) return toast("Lieferung nicht gefunden – bitte Aktualisieren.", true);
    var lm = lineMap(lines());
    openModal(modalShell("Eingang " + s.lieferid, "receive",
      '<input type="hidden" name="lieferid" value="' + esc(s.lieferid) + '">' +
      '<p class="muted">' + esc(shipArtikel(s, lm) || "") + (s.bestellnr ? " · Bestellung " + esc(s.bestellnr + " / " + s.pos) : " · ohne Bestellung") +
      (S.hasNum(s.menge_versendet) ? " · versendet " + qty(s.menge_versendet) : "") + "</p>" + receiveFields(s)));
  }

  function receiveFields(s) {
    s = s || {};
    var mErh = S.hasNum(s.menge_erhalten) ? s.menge_erhalten : (S.hasNum(s.menge_versendet) ? s.menge_versendet : "");
    return '<div class="row2">' + field("Eingangsdatum", '<input name="eingangsdatum" type="date" required value="' + esc(s.eingangsdatum || today()) + '">') +
      field("Menge erhalten (gezählt)", '<input name="menge_erhalten" type="number" min="0" step="any" inputmode="decimal" required value="' + esc(mErh) + '">') + "</div>" +
      field("QS-Ergebnis (leer = QS folgt)", seg("qs_ergebnis", SCHEMA.lists.qs_ergebnis, s.qs_ergebnis || "")) +
      '<div class="row2">' + field("Menge i.O.", '<input name="menge_io" type="number" min="0" step="any" inputmode="decimal" value="' + esc(S.hasNum(s.menge_io) ? s.menge_io : "") + '">') +
      field("QS geprüft von", '<input name="qs_von" value="' + esc(s.qs_von || userName()) + '">') + "</div>" +
      field("QS-Bemerkung", '<input name="qs_bemerkung" value="' + esc(s.qs_bemerkung || "") + '">');
  }

  function modalReceiveNew() {
    var ls = lines();
    openModal(modalShell("Paket ohne Versandmeldung", "receiveNew",
      field("Richtung", seg("richtung", SCHEMA.lists.richtung, S.DIR_BH)) +
      field("Bestellposition", '<select name="line"><option value="-">Ohne Bestellung</option>' + openLineOptions(ls) + "</select>") +
      field("Artikel (wenn ohne Bestellung)", '<input name="artikel" list="stocklist">') +
      field("Grund / Absender / Lieferschein", '<input name="bemerkung">') +
      '<label class="chk f"><input type="checkbox" name="express"> Express</label>' +
      receiveFields(null) + stockDatalist()));
  }

  function modalProblem(p, bezug, kat) {
    var isNew = !p; p = p || {};
    var L = SCHEMA.lists;
    openModal(modalShell(isNew ? "Neues Problem" : "Problem #" + p.nr, "problem",
      '<input type="hidden" name="nr" value="' + esc(p.nr || "") + '">' +
      field("Problem", '<textarea name="problem" required>' + esc(p.problem || "") + "</textarea>") +
      '<div class="row2">' + field("Kategorie", '<select name="kategorie">' + options(L.kategorie, p.kategorie || kat || "Sonstiges") + "</select>") +
      field("Bezug (BestellNr / LieferID)", '<input name="bezug" value="' + esc(p.bezug || bezug || "") + '">') + "</div>" +
      '<div class="row2">' + field("Verantwortlich", '<input name="verantwortlich" required value="' + esc(p.verantwortlich || "") + '">') +
      field("Standort", '<select name="standort">' + options(L.standort, p.standort || "Berlin") + "</select>") + "</div>" +
      '<div class="row2">' + field("Dringlichkeit", '<select name="dringlichkeit">' + options(L.dringlichkeit, p.dringlichkeit || "mittel") + "</select>") +
      field("Termin", '<input name="termin" type="date" value="' + esc(p.termin || "") + '">') + "</div>" +
      field("Maßnahme (sofort / langfristig)", '<textarea name="massnahme">' + esc(p.massnahme || "") + "</textarea>") +
      field("Status", '<select name="status">' + options(L.status, p.status || "offen") + "</select>")));
  }

  function afterSave(msg) { closeModal(); busy(false); toast(msg); return loadAll(); }

  var FORMS = {
    newOrder: function (f, d) {
      var snr = f.querySelectorAll("[name=ol_snr]"), art = f.querySelectorAll("[name=ol_artikel]"), men = f.querySelectorAll("[name=ol_menge]");
      var pos = [];
      for (var i = 0; i < snr.length; i++) {
        var sn = snr[i].value.trim(), ar = art[i].value.trim(), mg = men[i].value.trim();
        if (!sn && !ar && !mg) continue;
        if (!ar || !S.hasNum(mg) || S.num(mg) <= 0) throw new Error("Position " + (i + 1) + ": Artikel und Menge > 0 erforderlich.");
        pos.push({ snr: sn, artikel: ar, menge: S.num(mg), einheit: D.stock[sn] ? D.stock[sn].unit : "" });
      }
      if (!pos.length) throw new Error("Mindestens eine Position erfassen.");
      var lb = lbValue(d);
      return readTable("orders").then(function (rows) {
        var same = rows.filter(function (r) { return r.bestellnr === d.bestellnr; });
        if (same.length && !confirm("BestellNr " + d.bestellnr + " existiert bereits. Positionen anhängen?")) throw new Error("Abgebrochen.");
        var max = 0;
        same.forEach(function (r) { max = Math.max(max, S.num(r.pos)); });
        return addRows("orders", pos.map(function (p, i) {
          return { bestellnr: d.bestellnr, pos: max + i + 1, snr: p.snr, artikel: p.artikel, menge: p.menge, einheit: p.einheit,
                   wunschtermin: d.wunschtermin, bestaetigt: "", express: d.express, abgeschlossen: false,
                   erstellt_von: userName(), erstellt_am: today(), bemerkung: d.bemerkung, lbnr: lb };
        }));
      }).then(function () { return afterSave("Bestellung " + d.bestellnr + " gespeichert (" + pos.length + " Pos.)"); });
    },
    editOrder: function (f, d) {
      if (!S.hasNum(d.menge) || S.num(d.menge) <= 0) throw new Error("Menge > 0 erforderlich.");
      var patch = { menge: S.num(d.menge), wunschtermin: d.wunschtermin, bestaetigt: d.bestaetigt, express: d.express,
                    abgeschlossen: d.abgeschlossen, bemerkung: d.bemerkung }, lb = lbValue(d);
      if (lb !== undefined) patch.lbnr = lb;
      return updateRow("orders", function (r) { return S.lineKey(r.bestellnr, r.pos) === d.key; }, patch)
        .then(function () { return afterSave("Gespeichert."); });
    },
    ship: function (f, d) {
      var bh = d.richtung === S.DIR_BH, row = { richtung: d.richtung, versanddatum: d.versanddatum, menge_versendet: S.num(d.menge),
        express: d.express, tracking: d.tracking, versendet_von: userName(), bemerkung: d.bemerkung };
      if (!S.hasNum(d.menge) || S.num(d.menge) <= 0) throw new Error("Menge > 0 erforderlich.");
      if (bh && !d.line) throw new Error("Bitte Bestellposition wählen (oder 'Ohne Bestellung').");
      if (bh && d.line !== "-") {
        var l = lines().filter(function (x) { return x.key === d.line; })[0];
        if (!l) throw new Error("Bestellposition nicht gefunden – bitte Aktualisieren.");
        row.bestellnr = l.bestellnr; row.pos = l.pos; row.artikel = l.artikel;
        if (S.num(d.menge) > l.open - l.transit && !confirm("Menge ist größer als die offene Menge (" + (l.open - l.transit) + "). Trotzdem melden?")) throw new Error("Abgebrochen.");
      } else {
        if (!d.artikel) throw new Error("Artikel angeben.");
        if (bh && !d.bemerkung) throw new Error("Ohne Bestellung: Grund/Bezug angeben.");
        row.artikel = d.artikel;
      }
      return saveShipment(row, true);
    },
    receive: function (f, d) {
      checkReceipt(d);
      return updateRow("shipments", function (r) { return r.lieferid === d.lieferid; }, receiptPatch(d))
        .then(function () { return afterSave("Eingang " + d.lieferid + " gespeichert."); })
        .then(function () { askProblem(d, "Lieferung " + d.lieferid); });
    },
    receiveNew: function (f, d) {
      checkReceipt(d);
      var row = receiptPatch(d);
      row.richtung = d.richtung; row.express = d.express; row.bemerkung = d.bemerkung;
      if (d.line && d.line !== "-") {
        var l = lines().filter(function (x) { return x.key === d.line; })[0];
        if (!l) throw new Error("Bestellposition nicht gefunden – bitte Aktualisieren.");
        row.bestellnr = l.bestellnr; row.pos = l.pos; row.artikel = l.artikel;
      } else {
        if (!d.artikel) throw new Error("Artikel angeben.");
        if (!d.bemerkung) throw new Error("Ohne Bestellung: Grund/Absender angeben.");
        row.artikel = d.artikel;
      }
      return saveShipment(row, false).then(function (id) { askProblem(d, "Lieferung " + id); });
    },
    problem: function (f, d) {
      if (!d.problem || !d.verantwortlich) throw new Error("Problem und Verantwortlich ausfüllen.");
      var patch = { kategorie: d.kategorie, bezug: d.bezug, problem: d.problem, verantwortlich: d.verantwortlich,
                    standort: d.standort, dringlichkeit: d.dringlichkeit, massnahme: d.massnahme, termin: d.termin, status: d.status };
      if (d.nr) {
        var old = D.problems.filter(function (p) { return String(p.nr) === d.nr; })[0];
        if (d.status === "erledigt" && (!old || old.status !== "erledigt")) patch.erledigt_am = today();
        if (d.status !== "erledigt") patch.erledigt_am = "";
        return updateRow("problems", function (r) { return String(r.nr) === d.nr; }, patch)
          .then(function () { return afterSave("Problem #" + d.nr + " gespeichert."); });
      }
      return readTable("problems").then(function (rows) {
        patch.nr = S.nextProblemNr(rows); patch.datum = today();
        if (d.status === "erledigt") patch.erledigt_am = today();
        return addRows("problems", [patch]).then(function () { return afterSave("Problem #" + patch.nr + " angelegt."); });
      });
    },
    saveMeeting: function (f, d) {
      var per = S.meetingPeriod(D.meetings, today());
      var ks = S.kpis(lines(), D.shipments, D.problems, per.from, per.to, today(), CFG.kpi_targets);
      var row = { datum: today(), uhrzeit: new Date().toLocaleTimeString("de-DE", { hour: "2-digit", minute: "2-digit" }),
                  teilnehmer: d.teilnehmer, dauer: Math.max(1, Math.round((Date.now() - M.start) / 60000)), notizen: d.notizen, erstellt_von: userName() };
      ks.forEach(function (k) { row[k.key] = k.pct ? S.fmtKpi(k) : k.value; });
      return addRows("meetings", [row]).then(function () {
        M = { step: 0, start: null, stepStart: null };
        return afterSave("Meeting gespeichert.");
      });
    }
  };

  function checkReceipt(d) {
    if (!S.hasNum(d.menge_erhalten)) throw new Error("Menge erhalten angeben.");
    if (d.qs_ergebnis && !d.qs_von) throw new Error("QS geprüft von angeben.");
    if (S.hasNum(d.menge_io) && S.num(d.menge_io) > S.num(d.menge_erhalten)) throw new Error("Menge i.O. größer als Menge erhalten.");
    if (d.qs_ergebnis === "teilweise" && !S.hasNum(d.menge_io)) throw new Error("Bei 'teilweise' die Menge i.O. angeben.");
  }
  function receiptPatch(d) {
    var io = d.menge_io;
    if (!S.hasNum(io) && d.qs_ergebnis === "i.O.") io = d.menge_erhalten;
    if (!S.hasNum(io) && d.qs_ergebnis === "n.i.O.") io = 0;
    return { eingangsdatum: d.eingangsdatum, menge_erhalten: S.num(d.menge_erhalten), qs_ergebnis: d.qs_ergebnis,
             menge_io: io, qs_von: d.qs_ergebnis ? d.qs_von : "", qs_bemerkung: d.qs_bemerkung };
  }
  function askProblem(d, bezug) {
    if ((d.qs_ergebnis === "n.i.O." || d.qs_ergebnis === "teilweise") && confirm("QS nicht i.O. – jetzt Problemkarte anlegen?")) {
      modalProblem(null, bezug, "Qualität");
    }
  }
  function saveShipment(row, showId) {
    return readTable("shipments").then(function (rows) {
      var day = row.versanddatum || row.eingangsdatum || today();
      row.lieferid = S.nextLieferId(rows, S.normDate(day) || today());
      return addRows("shipments", [row]);
    }).then(function () {
      busy(false);
      if (showId) {
        openModal("<h2>Versand gemeldet</h2><p>Bitte auf Lieferschein und Paket schreiben:</p>" +
          '<div class="bigid">' + esc(row.lieferid) + "</div>" +
          (row.bestellnr ? '<div class="bigid" style="font-size:22px">Bestellung ' + esc(row.bestellnr + " / " + row.pos) + "</div>" : "") +
          '<div class="bar"><button class="btn pri" type="button" data-act="close">OK</button></div>');
        toast("Versand " + row.lieferid + " gespeichert.");
        loadAll();
      } else {
        closeModal(); toast("Eingang " + row.lieferid + " gespeichert."); loadAll();
      }
      return row.lieferid;
    });
  }

  /* ---------- events ---------- */

  var ACT = {
    close: closeModal,
    mstep: function (el) {
      var i = parseInt(el.getAttribute("data-i"), 10);
      if (i < 0 || i >= STEPS.length) return;
      M.step = i; M.stepStart = Date.now(); render(); window.scrollTo(0, 0);
    },
    mreset: function () { M = { step: 0, start: null, stepStart: null }; render(); },
    seg: function (el) {
      var f = el.closest("form"), name = el.getAttribute("data-name"), val = el.getAttribute("data-val");
      var inp = f.querySelector('input[type=hidden][name="' + name + '"]');
      if (inp.value === val && name === "qs_ergebnis") val = "";
      inp.value = val;
      Array.prototype.forEach.call(el.parentNode.children, function (b) { b.className = b.getAttribute("data-val") === val ? "on" : ""; });
      if (name === "richtung") syncShipForm(f);
      if (name === "qs_ergebnis") {
        var io = f.querySelector("[name=menge_io]"), me = f.querySelector("[name=menge_erhalten]");
        if (val === "i.O." && io && !io.value) io.value = me.value;
        if (val === "n.i.O." && io && !io.value) io.value = "0";
      }
    },
    ofilter: function (el) { filt.orders = el.getAttribute("data-f"); render(); },
    pfilter: function (el) { filt.problems = el.getAttribute("data-f"); render(); },
    newOrder: modalNewOrder,
    addLine: function () { $("olines").insertAdjacentHTML("beforeend", orderLineRow()); },
    rmLine: function (el) { var p = el.parentNode; if (p.parentNode.children.length > 1) p.parentNode.removeChild(p); },
    editOrder: function (el) { modalEditOrder(el.getAttribute("data-k")); },
    receive: function (el) { modalReceive(el.getAttribute("data-id")); },
    receiveNew: modalReceiveNew,
    newProblem: function (el) { modalProblem(null, el.getAttribute("data-bezug"), el.getAttribute("data-kat")); },
    editProblem: function (el) {
      var p = D.problems.filter(function (x) { return String(x.nr) === el.getAttribute("data-nr"); })[0];
      if (p) modalProblem(p);
    },
    probStatus: function (el) {
      var nr = el.getAttribute("data-nr"), s = el.getAttribute("data-s"), patch = { status: s };
      if (s === "erledigt") patch.erledigt_am = today();
      busy(true);
      updateRow("problems", function (r) { return String(r.nr) === nr; }, patch)
        .then(function () { busy(false); toast("Problem #" + nr + ": " + s); return loadAll(); })["catch"](fail);
    }
  };

  function syncShipForm(f) {
    var dir = f.querySelector("[name=richtung]").value, line = f.querySelector("[name=line]");
    var bh = dir === S.DIR_BH, bhBox = f.querySelector("[data-dir=bh]"), free = f.querySelector("[data-dir=free]");
    if (!bhBox) return;
    bhBox.style.display = bh ? "" : "none";
    free.style.display = !bh || (line && line.value === "-") ? "" : "none";
  }

  function bindEvents() {
    document.addEventListener("click", function (e) {
      var el = e.target.closest ? e.target.closest("[data-act]") : null;
      if (el && ACT[el.getAttribute("data-act")]) {
        e.preventDefault();
        try { ACT[el.getAttribute("data-act")](el, e); } catch (x) { fail(x); }
      } else if (e.target === $("overlay")) {
        /* clicks on the dark background do nothing, to avoid losing input */
      }
    });
    document.addEventListener("submit", function (e) {
      var f = e.target, name = f.getAttribute("data-form");
      if (!name || !FORMS[name]) return;
      e.preventDefault();
      var btn = f.querySelector("[type=submit]");
      if (btn) btn.disabled = true;
      busy(true);
      var done = function () { if (btn) btn.disabled = false; };
      try {
        Promise.resolve(FORMS[name](f, formData(f))).then(done, function (x) { done(); fail(x); });
      } catch (x) { done(); fail(x); }
    });
    document.addEventListener("change", function (e) {
      var t = e.target;
      if (t.name === "line") syncShipForm(t.closest("form"));
      if (t.getAttribute && t.getAttribute("data-snr")) {
        var s = D.stock[t.value.trim()], art = t.parentNode.querySelector("[name=ol_artikel]");
        if (s && art && !art.value) art.value = s.name;
      }
    });
    document.addEventListener("input", function (e) {
      if (e.target.id === "oq") { filt.q = e.target.value; $("olist").innerHTML = orderList(lines()); }
    });
    window.addEventListener("hashchange", function () { setView(); render(); });
    $("reload").onclick = function () { loadAll(); };
    $("logout").onclick = function () { msalApp.logoutRedirect({ account: account }); };
    setInterval(tick, 1000);
    setInterval(function () {
      if (view === "meeting" && M.step !== STEPS.length - 1 && $("overlay").style.display !== "flex" && !document.hidden) loadAll(true);
    }, (CFG.refresh_seconds || 60) * 1000);
  }

  function setView() {
    var h = (location.hash || "").replace(/^#/, "");
    view = ["meeting", "bestellungen", "versand", "eingang", "probleme"].indexOf(h) >= 0 ? h : "meeting";
  }

  /* ---------- start ---------- */

  function start() {
    $("login").style.display = "none";
    $("app").style.display = "block";
    $("info").textContent = userName();
    if (!CFG.shopfloor_share_link) { fail(new Error("config.json: shopfloor_share_link fehlt.")); return; }
    bindEvents();
    setView();
    loadAll();
  }

  function init(res) {
    CFG = res[0]; SCHEMA = res[1];
    document.title = CFG.title || document.title;
    $("title").textContent = CFG.title || "Shopfloor";
    if (!CFG.client_id || !CFG.tenant_id) throw new Error("config.json: client_id / tenant_id fehlt.");
    msalApp = new msal.PublicClientApplication({
      auth: { clientId: CFG.client_id, authority: CFG.authority_base.replace(/\/$/, "") + "/" + CFG.tenant_id, redirectUri: redirectUri() },
      cache: { cacheLocation: "localStorage", storeAuthStateInCookie: true }
    });
    return msalApp.handleRedirectPromise().then(function (resp) {
      if (resp && resp.account) account = resp.account;
      else { var a = msalApp.getAllAccounts(); account = a.length ? a[0] : null; }
      if (!account) {
        $("login").style.display = "block";
        $("loginbtn").onclick = function () { msalApp.loginRedirect({ scopes: CFG.scopes }); };
        return;
      }
      msalApp.setActiveAccount(account);
      start();
    });
  }

  Promise.all([getJSON("config.json"), getJSON("schema.json")]).then(init)["catch"](function (e) {
    $("login").style.display = "block";
    $("loginbtn").style.display = "none";
    $("lerr").textContent = String(e && e.message ? e.message : e);
  });
})();
