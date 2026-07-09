function doPost(e) {
  var sheet = getOrdersSheet();
  var data = JSON.parse(e.postData.contents);
  sheet.appendRow([
    new Date(data.time),
    data.customerType,
    data.dining,
    data.items.map(function(it){ return it.qty + "x " + it.name; }).join(", "),
    data.subtotal / 100,
    data.surcharge / 100,
    data.total / 100
  ]);
  return ContentService.createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
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
  return ContentService.createTextOutput(JSON.stringify({ orders: orders }))
    .setMimeType(ContentService.MimeType.JSON);
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
