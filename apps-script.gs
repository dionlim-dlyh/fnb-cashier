function doPost(e) {
  var data = JSON.parse(e.postData.contents);
  if (data.type === "dailysales") {
    return saveDailySales(data);
  }
  return saveOrder(data);
}

function doGet(e) {
  if (e.parameter.source === "dailysales") {
    return getDailySales();
  }
  return getOrders();
}

function saveOrder(data){
  var sheet = getOrdersSheet();
  sheet.appendRow([
    new Date(data.time),
    data.customerType,
    data.dining,
    data.items.map(function(it){ return it.qty + "x " + it.name; }).join(", "),
    data.subtotal / 100,
    data.surcharge / 100,
    data.total / 100
  ]);
  return jsonOutput({ ok: true });
}

function getOrders(){
  var sheet = getOrdersSheet();
  var rows = sheet.getDataRange().getValues();
  rows.shift();
  var orders = rows
    .filter(function(r){ return r[0]; })
    .map(function(r){
      return {
        time: r[0] instanceof Date ? r[0].toISOString() : r[0],
        customerType: r[1],
        dining: r[2],
        items: r[3],
        subtotal: Number(r[4]) || 0,
        surcharge: Number(r[5]) || 0,
        total: Number(r[6]) || 0
      };
    });
  return jsonOutput({ orders: orders });
}

function saveDailySales(data){
  var sheet = getDailySalesSheet();
  var dateStr = data.date;
  var dateObj = new Date(dateStr + "T00:00:00");
  var day = dateObj.toLocaleDateString("en-US", { weekday: "short" });
  var payNow = Number(data.payNow) || 0;
  var cash = Number(data.cash) || 0;
  var closed = !!data.closed;
  var remarks = data.remarks || "";

  var values = sheet.getDataRange().getValues();
  var rowIndex = -1;
  for (var i = 1; i < values.length; i++) {
    var existing = values[i][0];
    var existingStr = existing instanceof Date ? Utilities.formatDate(existing, Session.getScriptTimeZone(), "yyyy-MM-dd") : existing;
    if (existingStr === dateStr) { rowIndex = i + 1; break; }
  }

  var row = [dateObj, day, payNow, cash, closed ? "Closed" : "", remarks];
  if (rowIndex > 0) {
    sheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
  } else {
    sheet.appendRow(row);
  }
  return jsonOutput({ ok: true });
}

function getDailySales(){
  var sheet = getDailySalesSheet();
  var rows = sheet.getDataRange().getValues();
  rows.shift();
  var entries = rows
    .filter(function(r){ return r[0]; })
    .map(function(r){
      return {
        date: r[0] instanceof Date ? Utilities.formatDate(r[0], Session.getScriptTimeZone(), "yyyy-MM-dd") : r[0],
        day: r[1],
        payNow: Number(r[2]) || 0,
        cash: Number(r[3]) || 0,
        closed: r[4] === "Closed",
        remarks: r[5] || ""
      };
    });
  return jsonOutput({ entries: entries });
}

function jsonOutput(obj){
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function getOrdersSheet(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Orders");
  if(!sheet){
    sheet = ss.insertSheet("Orders");
    sheet.appendRow(["Time","Customer Type","Dining","Items","Subtotal","Surcharge","Total"]);
  }
  return sheet;
}

function getDailySalesSheet(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Daily Sales");
  if(!sheet){
    sheet = ss.insertSheet("Daily Sales");
    sheet.appendRow(["Date","Day","PayNow","Cash","Closed","Remarks"]);
  }
  return sheet;
}
