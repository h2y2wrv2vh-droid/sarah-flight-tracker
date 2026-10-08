/* ---------- flexible dates + deal info (v6) ---------- */

function shiftDate(iso, days) {
  var d = new Date(String(iso).substring(0, 10) + 'T12:00:00');
  d.setDate(d.getDate() + days);
  return Utilities.formatDate(d, 'America/Toronto', 'yyyy-MM-dd');
}

// Cheapest price for departDate +/- 3 days (round trips keep the same trip length).
// NOTE: up to 7 live SerpApi searches — call only on explicit user tap.
function flexDates(p) {
  var destCode = String(p.destination || '').trim().toUpperCase();
  if (!/^[A-Z]{3}$/.test(destCode)) return {ok: false, error: 'airport code required'};
  var origin = (String(p.homeAirport || 'YYZ').trim() || 'YYZ').toUpperCase();
  var dep = String(p.departDate || '').trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(dep)) return {ok: false, error: 'departDate required'};
  var ret = String(p.returnDate || '').trim();
  var tripLen = 0;
  if (/^\d{4}-\d{2}-\d{2}$/.test(ret)) {
    tripLen = Math.round((new Date(ret + 'T12:00:00') - new Date(dep + 'T12:00:00')) / 86400000);
  }
  var out = [];
  for (var d = -3; d <= 3; d++) {
    var dd = shiftDate(dep, d);
    var rd = tripLen > 0 ? shiftDate(dep, d + tripLen) : '';
    try {
      var t = buildTrip({
        destination: destCode, departDate: dd, returnDate: rd,
        homeAirport: origin, travelers: p.travelers, cabin: p.cabin,
        nonstop: p.nonstop === '1' || p.nonstop === 'true'
      });
      var res = findOffers(t);
      var minP = null;
      if (res && res.offers) {
        for (var i = 0; i < res.offers.length; i++) {
          var pr = parseFloat(res.offers[i].price);
          if (!isNaN(pr) && (minP === null || pr < minP)) minP = pr;
        }
      }
      out.push({date: dd, returnDate: rd, minPrice: minP, ok: !!(res && res.offers && res.offers.length)});
    } catch (e) {
      out.push({date: dd, returnDate: rd, minPrice: null, ok: false});
    }
  }
  return {ok: true, origin: origin, dest: destCode, days: out};
}

// 30-day route price stats from PriceHistory (joined via Trips).
function dealInfo(p) {
  var origin = (String(p.origin || p.homeAirport || '').trim() || '').toUpperCase();
  var dest = (String(p.dest || p.destination || '').trim() || '').toUpperCase();
  if (!origin || !/^[A-Z]{3}$/.test(dest)) return {ok: true, count: 0};
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var tmap = {};
  var trips = ss.getSheetByName('Trips');
  if (trips && trips.getLastRow() > 1) {
    var rows = trips.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      var tid = String(rows[i][0] || '');
      if (tid && String(rows[i][2] || '').toUpperCase() === dest &&
          String(rows[i][3] || '').toUpperCase() === origin) tmap[tid] = 1;
    }
  }
  var prices = [];
  var ph = ss.getSheetByName('PriceHistory');
  if (ph && ph.getLastRow() > 1) {
    var cutoff = new Date();
    cutoff.setDate(cutoff.getDate() - 30);
    var prow = ph.getDataRange().getValues();
    for (var j = 1; j < prow.length; j++) {
      if (!tmap[String(prow[j][1] || '')]) continue;
      var pr = parseFloat(prow[j][2]);
      if (isNaN(pr)) continue;
      var dt = new Date(String(prow[j][0] || '') + 'T12:00:00');
      if (isNaN(dt.getTime()) || dt < cutoff) continue;
      prices.push(pr);
    }
  }
  if (prices.length < 3) return {ok: true, count: prices.length};
  var sum = 0, min = prices[0];
  for (var k = 0; k < prices.length; k++) { sum += prices[k]; if (prices[k] < min) min = prices[k]; }
  return {ok: true, count: prices.length, avg: Math.round(sum / prices.length), min: min};
}
