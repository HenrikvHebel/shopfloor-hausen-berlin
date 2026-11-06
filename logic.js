/* Shopfloor Hausen-Berlin: pure status and KPI rules (no DOM, no network). */
(function (root) {
  "use strict";

  var DIR_BH = "Berlin → Hausen";

  function pad(n) { return (n < 10 ? "0" : "") + n; }

  function isoOf(d) { return d.getFullYear() + "-" + pad(d.getMonth() + 1) + "-" + pad(d.getDate()); }

  /* Excel cell value -> "YYYY-MM-DD" or "" (accepts serial numbers, ISO and dd.mm.yyyy). */
  function normDate(v) {
    if (v === null || v === undefined || v === "") return "";
    if (typeof v === "number") {
      if (v < 1) return "";
      var d = new Date(Date.UTC(1899, 11, 30) + Math.round(v * 86400000));
      return d.getUTCFullYear() + "-" + pad(d.getUTCMonth() + 1) + "-" + pad(d.getUTCDate());
    }
    var s = String(v).trim(), m;
    if ((m = /^(\d{4})-(\d{1,2})-(\d{1,2})/.exec(s))) return m[1] + "-" + pad(+m[2]) + "-" + pad(+m[3]);
    if ((m = /^(\d{1,2})\.(\d{1,2})\.(\d{2,4})/.exec(s))) {
      var y = +m[3]; if (y < 100) y += 2000;
      return y + "-" + pad(+m[2]) + "-" + pad(+m[1]);
    }
    if (/^\d+(\.\d+)?$/.test(s)) return normDate(parseFloat(s));
    return "";
  }

  function fmtDate(iso) {
    if (!iso) return "";
    var p = iso.split("-");
    return p[2] + "." + p[1] + "." + p[0];
  }

  function num(v) {
    if (v === null || v === undefined || v === "") return 0;
    var n = typeof v === "number" ? v : parseFloat(String(v).replace(",", "."));
    return isNaN(n) ? 0 : n;
  }

  function hasNum(v) { return v !== null && v !== undefined && String(v).trim() !== "" && !isNaN(num(v)); }

  function yes(v) { return /^(j|ja|y|yes|x|1|true|wahr)$/i.test(String(v === undefined || v === null ? "" : v).trim()); }

  function str(v) { return v === null || v === undefined ? "" : String(v).trim(); }

  function lineKey(bestellnr, pos) { return str(bestellnr) + "|" + str(num(pos)); }

  function daysBetween(a, b) {
    if (!a || !b) return 0;
    return Math.round((Date.parse(b + "T00:00:00Z") - Date.parse(a + "T00:00:00Z")) / 86400000);
  }

  function addDays(iso, n) {
    var d = new Date(Date.parse(iso + "T00:00:00Z") + n * 86400000);
    return d.getUTCFullYear() + "-" + pad(d.getUTCMonth() + 1) + "-" + pad(d.getUTCDate());
  }

  function isReceived(s) { return !!normDate(s.eingangsdatum); }

  /* Quantity that passed QS for a received shipment. */
  function okQty(s) {
    if (!isReceived(s)) return 0;
    var r = str(s.qs_ergebnis);
    if (hasNum(s.menge_io)) return num(s.menge_io);
    if (r === "i.O.") return num(s.menge_erhalten);
    return 0;
  }

  /*
   * Order lines with derived state.
   * status: Abgeschlossen | Geliefert | Unterwegs | Teilgeliefert | Offen
   * flags:  Überfällig, Verspätet, Mengenabweichung, QS n.i.O., QS offen, Express
   */
  function deriveLines(orders, shipments, today) {
    var byKey = {};
    (shipments || []).forEach(function (s) {
      if (!str(s.bestellnr)) return;
      var k = lineKey(s.bestellnr, s.pos);
      (byKey[k] = byKey[k] || []).push(s);
    });
    return (orders || []).filter(function (o) { return str(o.bestellnr); }).map(function (o) {
      var k = lineKey(o.bestellnr, o.pos), ss = byKey[k] || [];
      var ordered = num(o.menge), shipped = 0, received = 0, ok = 0, transit = 0, lastIn = "";
      var qsFail = false, qsOpen = false, diff = false;
      ss.forEach(function (s) {
        if (str(s.richtung) && str(s.richtung) !== DIR_BH) return;
        shipped += num(s.menge_versendet);
        if (isReceived(s)) {
          received += num(s.menge_erhalten);
          ok += okQty(s);
          var din = normDate(s.eingangsdatum);
          if (din > lastIn) lastIn = din;
          var r = str(s.qs_ergebnis);
          if (r === "n.i.O." || r === "teilweise") qsFail = true;
          if (!r) qsOpen = true;
          if (hasNum(s.menge_versendet) && num(s.menge_versendet) !== num(s.menge_erhalten)) diff = true;
        } else {
          transit += num(s.menge_versendet);
        }
      });
      var due = normDate(o.bestaetigt) || normDate(o.wunschtermin);
      var status;
      if (yes(o.abgeschlossen)) status = "Abgeschlossen";
      else if (ordered > 0 && received >= ordered) status = "Geliefert";
      else if (transit > 0) status = "Unterwegs";
      else if (received > 0) status = "Teilgeliefert";
      else status = "Offen";
      var done = status === "Geliefert" || status === "Abgeschlossen";
      if (received > ordered) diff = true;
      var flags = [];
      if (!done && due && due < today) flags.push("Überfällig");
      if (done && due && lastIn && lastIn > due) flags.push("Verspätet");
      if (diff) flags.push("Mengenabweichung");
      if (qsFail) flags.push("QS n.i.O.");
      if (qsOpen) flags.push("QS offen");
      if (yes(o.express)) flags.push("Express");
      return {
        key: k, order: o, bestellnr: str(o.bestellnr), pos: num(o.pos), snr: str(o.snr), artikel: str(o.artikel),
        einheit: str(o.einheit), ordered: ordered, shipped: shipped, received: received, okQty: ok,
        transit: transit, open: Math.max(0, ordered - received), due: due, lastIn: lastIn,
        status: status, done: done, flags: flags, shipments: ss
      };
    });
  }

  function hasFlag(line, f) { return line.flags.indexOf(f) >= 0; }

  /* Sort for meeting: overdue first, then by due date, then order number. */
  function sortLines(lines) {
    return lines.slice().sort(function (a, b) {
      var ao = hasFlag(a, "Überfällig") ? 0 : 1, bo = hasFlag(b, "Überfällig") ? 0 : 1;
      if (ao !== bo) return ao - bo;
      var ad = a.due || "9999", bd = b.due || "9999";
      if (ad !== bd) return ad < bd ? -1 : 1;
      if (a.bestellnr !== b.bestellnr) return a.bestellnr < b.bestellnr ? -1 : 1;
      return a.pos - b.pos;
    });
  }

  function ratio(n, d) { return d ? n / d : null; }

  /*
   * KPIs for the period [from, to] (ISO dates, inclusive).
   * targets: { otd_min, qty_min, qs_min, with_order_min, qs_open_max, overdue_max, problems_open_max }
   */
  function kpis(lines, shipments, problems, from, to, today, targets) {
    var t = targets || {};
    var lineMap = {};
    lines.forEach(function (l) { lineMap[l.key] = l; });
    var inPeriod = (shipments || []).filter(function (s) {
      var d = normDate(s.eingangsdatum);
      return d && d >= from && d <= to && (!str(s.richtung) || str(s.richtung) === DIR_BH);
    });
    var otdN = 0, otdD = 0, qtyN = 0, qtyD = 0, qsN = 0, qsD = 0, qsOpen = 0, withOrder = 0;
    inPeriod.forEach(function (s) {
      var l = str(s.bestellnr) ? lineMap[lineKey(s.bestellnr, s.pos)] : null;
      if (str(s.bestellnr)) withOrder++;
      if (l && l.due) { otdD++; if (normDate(s.eingangsdatum) <= l.due) otdN++; }
      if (hasNum(s.menge_versendet)) { qtyD++; if (num(s.menge_versendet) === num(s.menge_erhalten)) qtyN++; }
      var r = str(s.qs_ergebnis);
      if (r) { qsD++; if (r === "i.O.") qsN++; } else qsOpen++;
    });
    var overdue = lines.filter(function (l) { return hasFlag(l, "Überfällig"); }).length;
    var openProb = (problems || []).filter(function (p) { return str(p.problem) && str(p.status) !== "erledigt"; }).length;

    function minK(key, label, value, target, hint) {
      return { key: key, label: label, value: value, pct: true, target: target, dir: "min",
               ok: value === null || target === undefined ? null : value >= target, hint: hint };
    }
    function maxK(key, label, value, target, hint) {
      return { key: key, label: label, value: value, pct: false, target: target, dir: "max",
               ok: target === undefined ? null : value <= target, hint: hint };
    }
    return [
      minK("kpi_otd", "Liefertreue", ratio(otdN, otdD), t.otd_min, otdN + " von " + otdD + " Eingängen pünktlich"),
      minK("kpi_menge", "Mengentreue", ratio(qtyN, qtyD), t.qty_min, qtyN + " von " + qtyD + " Eingängen Menge wie versendet"),
      minK("kpi_qs", "QS-Quote", ratio(qsN, qsD), t.qs_min, qsN + " von " + qsD + " geprüften Eingängen i.O."),
      minK("kpi_bezug", "Mit Bestellbezug", ratio(withOrder, inPeriod.length), t.with_order_min,
           withOrder + " von " + inPeriod.length + " Eingängen mit BestellNr"),
      maxK("kpi_qs_offen", "Eingänge ohne QS", qsOpen, t.qs_open_max, "Eingänge ohne QS-Ergebnis"),
      maxK("kpi_ueberfaellig", "Überfällige Positionen", overdue, t.overdue_max, "offene Positionen nach Termin"),
      maxK("kpi_probleme", "Offene Probleme", openProb, t.problems_open_max, "Problemkarten nicht erledigt")
    ];
  }

  function fmtKpi(k) {
    if (k.value === null) return "–";
    return k.pct ? Math.round(k.value * 100) + " %" : String(k.value);
  }

  /* Next delivery ID for a date: L-YYYYMMDD-NN */
  function nextLieferId(shipments, iso) {
    var base = "L-" + iso.replace(/-/g, "") + "-", max = 0;
    (shipments || []).forEach(function (s) {
      var id = str(s.lieferid);
      if (id.indexOf(base) === 0) max = Math.max(max, parseInt(id.slice(base.length), 10) || 0);
    });
    return base + pad(max + 1);
  }

  function nextProblemNr(problems) {
    var max = 0;
    (problems || []).forEach(function (p) { max = Math.max(max, num(p.nr)); });
    return max + 1;
  }

  /* Meeting period: from the day of the last saved meeting (or 7 days back) until today. */
  function meetingPeriod(meetings, today) {
    var last = "";
    (meetings || []).forEach(function (m) {
      var d = normDate(m.datum);
      if (d && d < today && d > last) last = d;
    });
    return { from: last || addDays(today, -7), to: today, last: last };
  }

  var api = {
    DIR_BH: DIR_BH, normDate: normDate, fmtDate: fmtDate, isoOf: isoOf, num: num, hasNum: hasNum, yes: yes,
    str: str, lineKey: lineKey, daysBetween: daysBetween, addDays: addDays, isReceived: isReceived,
    okQty: okQty, deriveLines: deriveLines, hasFlag: hasFlag, sortLines: sortLines, kpis: kpis,
    fmtKpi: fmtKpi, nextLieferId: nextLieferId, nextProblemNr: nextProblemNr, meetingPeriod: meetingPeriod
  };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  root.SFL = api;
})(typeof window !== "undefined" ? window : this);
