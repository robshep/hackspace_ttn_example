// 2019 enhanced by Rob Shepherd @95point2, Jo Hinchliffe @concreted0g 
//  - auto create tab based on TTN device ID
//  - remove variable length gateways, just concat \n seperate gateways info parts
//  - use TTN payload fields in headers and data columns

// 2017 by Daniel Eichhorn, https://blog.squix.org
// Inspired by https://gist.github.com/bmcbride/7069aebd643944c9ee8b


// 1. Create or open an existing Sheet and click Tools > Script editor and enter the code below

// 2. Run > setup
// 3. Publish > Deploy as web app
//    - enter Project Version name and click 'Save New Version'
//    - set security level and enable service (most likely execute as 'me' and access 'anyone, even anonymously)
// 4. Copy the 'Current web app URL' and post this in your form/script action

var SCRIPT_PROP = PropertiesService.getScriptProperties(); // new property service


function doPost(e){
  var lock = LockService.getPublicLock();
  lock.waitLock(30000); // wait 30 seconds before conceding defeat.
  
  try {
    var jsonData = JSON.parse(e.postData.contents);
    // next set where we write the data - you could write to multiple/alternate destinations
    var doc = SpreadsheetApp.openById(SCRIPT_PROP.getProperty("key"));
    var sheet = doc.getSheetByName(jsonData.dev_id); 
    
    if(sheet == null){
      doc.insertSheet(jsonData.dev_id)
      sheet = doc.getSheetByName(jsonData.dev_id); 
    }
    
    // we'll assume header is in row 1 but you can override with header_row in GET/POST data
    //var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    var headerRow = [];
    // loop through the header columns
    
    // Update headers
    headerRow.push("App ID");
    headerRow.push("Device ID");
    headerRow.push("Device EUI");
    headerRow.push("Port");
    headerRow.push("Frame Counter");
    headerRow.push("Time");
    headerRow.push("Data Rate");
    // add headers for payload fields
    for (key in jsonData.payload_fields) {
      var value = jsonData.payload_fields[key];
      if (jsonData.payload_fields.hasOwnProperty(key)) {
        headerRow.push(key); 
      }
    }
    
    headerRow.push("Num Gateways");
    headerRow.push("Gateway Info");
    
    // set headers, replacing existing
    sheet.getRange(1, 1, 1, headerRow.length).setValues([headerRow]);
    
    // add new data row
    var row = [];
    row.push(jsonData.app_id);
    row.push(jsonData.dev_id);
    row.push(jsonData.hardware_serial);
    row.push(jsonData.port);
    row.push(jsonData.counter);
    row.push(jsonData.metadata.time);
    row.push(jsonData.metadata.data_rate);
    // add payload field data
    for (key in jsonData.payload_fields) {
      var value = jsonData.payload_fields[key];
      if (jsonData.payload_fields.hasOwnProperty(key)) {
        row.push(value); 
      }
    }
    
    
    row.push( jsonData.metadata.gateways.length );
    
    // make a gateway info string
    var gw_info = ""

    for (var i = 0; i < jsonData.metadata.gateways.length; i++) {
      var gateway = jsonData.metadata.gateways[i];
      gw_info += gateway.gtw_id;
      gw_info += ","
      gw_info += "rssi=" + gateway.rssi;
      gw_info += ","
      gw_info += "snr=" + gateway.snr;
      gw_info += ","
      gw_info += "loc=" + gateway.latitude + "," + gateway.longitude;

      if( i < jsonData.metadata.gateways.length-1){
        gw_info += "\n"
      }
    }
    
    row.push(gw_info);
    
    
    // set data row
    // more efficient to set values as [][] array than individually
    var nextRow = sheet.getLastRow()+1; // get next row
    sheet.getRange(nextRow, 1, 1, row.length).setValues([row]);
    // return json success results
    return ContentService
          .createTextOutput(JSON.stringify({"result":"success", "row": nextRow}))
          .setMimeType(ContentService.MimeType.JSON);
  } catch(e) {
    // if error return this
    return ContentService
          .createTextOutput(JSON.stringify({"result":"error", "error": e}))
          .setMimeType(ContentService.MimeType.JSON);
  } finally { //release lock
    lock.releaseLock();
  }
}

function setup() {
  var doc = SpreadsheetApp.getActiveSpreadsheet();
  SCRIPT_PROP.setProperty("key", doc.getId());
}
