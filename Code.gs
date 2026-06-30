var SPREADSHEET_ID = '';
var SALES_SHEET_NAME = 'SalesData';
var BRANCHES = [
  'Gorakhpur', 'Kolkata', 'Delhi', 'Kanpur', 'Surat', 'Varanasi',
  'Indore', 'Patna', 'Bhilwara', 'Haridwar', 'Rishikesh', 'Hyderabad',
  'Mumbai', 'Bengaluru', 'Chennai', 'Cuttack', 'Raipur', 'Jalgaon',
  'Nagpur', 'Ranchi'
];

function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'save') {
    var getPayload = JSON.parse(e.parameter.data || '{}');
    return jsonResponse_(saveSalesData(getPayload.sellerDetails, getPayload.cartItems));
  }

  var template = HtmlService.createTemplateFromFile('index');
  return template.evaluate()
    .setTitle('Book Sales Collection')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function doPost(e) {
  try {
    var contents = (e && e.postData && e.postData.contents) || '';
    var payloadText = (e && e.parameter && e.parameter.payload) || contents || '{}';
    var payload = JSON.parse(payloadText);
    return jsonResponse_(saveSalesData(payload.sellerDetails, payload.cartItems));
  } catch (err) {
    return jsonResponse_({ success: false, message: 'Invalid save request: ' + err.message });
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getBranches() {
  return BRANCHES.slice();
}

function getBookByCode(code) {
  return {
    success: false,
    message: 'Book code "' + code + '" was not found in the loaded catalog.'
  };
}

function saveSalesData(sellerDetails, cartItems) {
  try {
    validateSavePayload_(sellerDetails, cartItems);

    var ss = getSalesSpreadsheet_();
    var sheet = ss.getSheetByName(SALES_SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SALES_SHEET_NAME);
    }

    ensureSalesHeader_(sheet);

    var timestamp = new Date();
    var rows = cartItems.map(function(item) {
      return [
        sellerDetails.sellingDate,
        sellerDetails.branchName,
        sellerDetails.sellerName,
        sellerDetails.sellerMobile,
        sellerDetails.remarks || '',
        item.bookCode,
        item.bookName,
        Number(item.rate) || 0,
        Number(item.bundlePacking) || 0,
        Number(item.unitQty) || 0,
        Number(item.bundleQty) || 0,
        Number(item.totalUnits) || 0,
        Number(item.totalAmount) || 0,
        timestamp
      ];
    });

    var lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 14).setValues(rows);
    } finally {
      lock.releaseLock();
    }

    return {
      success: true,
      message: rows.length + ' item(s) saved successfully to "' + SALES_SHEET_NAME + '".'
    };
  } catch (err) {
    return { success: false, message: 'Error saving data: ' + err.message };
  }
}

function getSalesSpreadsheet_() {
  if (SPREADSHEET_ID) {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  }

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (spreadsheet) {
    return spreadsheet;
  }

  throw new Error('No spreadsheet is attached. Bind this Apps Script project to the target Google Sheet, or set SPREADSHEET_ID at the top of Code.gs.');
}

function ensureSalesHeader_(sheet) {
  var headers = [
    'Selling Date', 'Branch Name', 'Seller Name', 'Seller Mobile', 'Remarks',
    'Book Code', 'Book Name', 'Rate', 'Bundle Packing Number',
    'Unit Quantity', 'Bundle Quantity', 'Total Units', 'Total Amount', 'Submitted On'
  ];

  var firstCell = sheet.getRange(1, 1).getValue();
  if (firstCell === '' || firstCell !== headers[0]) {
    if (firstCell !== '') {
      sheet.insertRowBefore(1);
    }
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
}

function validateSavePayload_(sellerDetails, cartItems) {
  if (!sellerDetails) {
    throw new Error('Seller details are missing.');
  }
  if (!sellerDetails.sellingDate) {
    throw new Error('Selling date is missing.');
  }
  if (!sellerDetails.branchName) {
    throw new Error('Branch name is missing.');
  }
  if (!sellerDetails.sellerName) {
    throw new Error('Seller name is missing.');
  }
  if (!/^\d{10}$/.test(String(sellerDetails.sellerMobile || ''))) {
    throw new Error('Seller mobile must be exactly 10 digits.');
  }
  if (!Array.isArray(cartItems) || cartItems.length === 0) {
    throw new Error('Cart is empty.');
  }
}

function jsonResponse_(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
