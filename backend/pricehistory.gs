/* ---------- price history sparkline data (v7) ---------- */

// Price history time series for a route (powers the sparkline on tracked trips).
// Joins PriceHistory -> Trips on tripId, filters by origin/dest IATA codes.
function priceHistory(p) {
  var origin = (String(p.origin || p.homeAirport || '').trim() || '').toUpperCase();
  var dest = (String(p.dest || p.destination || '').trim() || '').toUpperCase();
  if (!origin || !/^[A-Z]{3}$/.test(dest)) return {ok: true, points: []};
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
  var byDay = {};
  var ph = ss.getSheetByName('PriceHistory');
  if (ph && ph.getLastRow() > 1) {
    var prow = ph.getDataRange().getValues();
    for (var j = 1; j < prow.length; j++) {
      if (!tmap[String(prow[j][1] || '')]) continue;
      var pr = parseFloat(prow[j][2]);
      if (isNaN(pr)) continue;
      var raw = prow[j][0];
      var ds = (raw instanceof Date)
        ? Utilities.formatDate(raw, 'UTC', 'yyyy-MM-dd')
        : String(raw || '').slice(0, 10);
      if (!/^\d{4}-\d{2}-\d{2}$/.test(ds)) continue;
      byDay[ds] = Math.round(pr);
    }
  }
  var days = Object.keys(byDay).sort();
  var points = [];
  for (var k = 0; k < days.length; k++) points.push({d: days[k], p: byDay[days[k]]});
  return {ok: true, points: points.slice(-90)};
}
