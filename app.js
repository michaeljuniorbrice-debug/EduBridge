const SPREADSHEET_ID = "13Gf68tCV8_Q835iWhVjn1TNONSepuFvXt_ymQRERVp4"; // YOUR SHEET ID

// ========== MAIN ROUTER ==========
function doGet(e) {
  const action = e.parameter.get;

  if (action == "checkUser") { return checkUser(e); } // Used by studentLogin
  if (action == "videos") { return getVideos(e); } // Used by loadContent
  if (action == "pdfs") { return getPDFs(e); } // Used by loadContent
  if (action == "allUsers") { return getAllUsers(e); }// Used by admin
  if (action == "resetDevice") { return resetDevice(e); } // Used by admin

  return jsonResponse({status: "error", message: "Invalid GET action"});
}

function doPost(e) {
  const data = JSON.parse(e.postData.contents);
  const type = data.type;

  if (type == "register") { return registerUser(data); } // Used by registerUser
  if (type == "lockDevice") { return lockDevice(data); } // Used by studentLogin
  if (type == "updateStatus") { return updateStatus(data); } // Used by admin
  if (type == "deleteUser") { return deleteUser(data); } // Used by admin
  if (type == "addPDF") { return addPDF(data); } // Used by admin
  if (type == "deletePDF") { return deletePDF(data); } // Used by admin
  if (type == "addVideo") { return addVideo(data); } // Used by admin
  if (type == "deleteVideo") { return deleteVideo(data); } // Used by admin

  return jsonResponse({status: "error", message: "Invalid POST type"});
}

// ========== 1. YOUR REGISTER FUNCTION - UNCHANGED LOGIC ==========
function registerUser(data){
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("Users");
  sheet.appendRow([
    new Date(), // Timestamp A
    "", // Level B - you can fill later
    data.Phone, // C
    data.Name, // D
    "", // Password E - not used
    "Pending", // F Status
    data.DeviceID || "", // G DeviceID
    data.PaidDate // H Expiry - we will use this as PaidDate. Expiry will be +30 days
  ]);
  return jsonResponse({status: "success"});
}

// ========== 2. YOUR LOGIN + DEVICE LOCK FUNCTION - UPDATED ==========
function checkUser(e){
  const phone = e.parameter.phone;
  const deviceID = e.parameter.device;

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("Users");
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    let rowPhone = data[i][2]; // C
    let status = data[i][5]; // F
    let savedDevice = data[i][6]; // G
    let expiry = data[i][7]; // H - Now we treat this as Expiry Date

    if (rowPhone == phone) {

      // 1. Check Status
      if (status!= "Active") {
        return jsonResponse({allow: false, reason: "not_active"});
      }

      // 2. Check Expiry
      if (expiry) {
        let today = new Date();
        let expiryDate = new Date(expiry);
        if (today > expiryDate) {
          return jsonResponse({allow: false, reason: "not_active"});
        }
      }

      // 3. DEVICE LOCK CHECK
      if (savedDevice && savedDevice!= "" && savedDevice!= deviceID) {
        return jsonResponse({allow: false, reason: "device_locked"});
      }

      // 4. ALLOWED
      return jsonResponse({allow: true});
    }
  }
  return jsonResponse({allow: false, reason: "not_found"});
}

// This is called after checkUser returns allow:true
function lockDevice(data){
  const phone = data.Phone;
  const deviceID = data.DeviceID;

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("Users");
  const range = sheet.getDataRange();
  const values = range.getValues();

  for (let i = 1; i < values.length; i++) {
    if (values[i][2] == phone) {
      sheet.getRange(i+1, 7).setValue(deviceID); // Save to Column G
      break;
    }
  }
  return jsonResponse({status: "success"});
}

// ========== 3. ADMIN FUNCTIONS ==========
function getAllUsers(e){
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("Users");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const users = [];

  for (let i = 1; i < data.length; i++) {
    users.push({
      Timestamp: data[i][0],
      Level: data[i][1],
      Phone: data[i][2],
      Name: data[i][3],
      Status: data[i][5],
      DeviceID: data[i][6],
      Expiry: data[i][7],
      TxID: "N/A", // You don't have TxID column. Add if needed
      PaidDate: data[i][7] // Using Expiry as PaidDate for now
    });
  }
  return jsonResponse({users: users});
}

function updateStatus(data){
  const txid = data.TxID; // We will use Phone instead since you don't have TxID
  const status = data.Status;
  const paidDate = data.PaidDate;
  const expiry = data.Expiry;

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("Users");
  const values = sheet.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {
    if (values[i][2] == txid) { // Matching by Phone
      sheet.getRange(i+1, 6).setValue(status); // F Status
      sheet.getRange(i+1, 8).setValue(expiry); // H Expiry
      break;
    }
  }
  return jsonResponse({status: "success"});
}

function deleteUser(data){
  const txid = data.TxID; // Matching by Phone
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("Users");
  const values = sheet.getDataRange().getValues();

  for (let i = values.length - 1; i >= 1; i--) {
    if (values[i][2] == txid) {
      sheet.deleteRow(i+1);
      break;
    }
  }
  return jsonResponse({status: "success"});
}

function resetDevice(e){
  const phone = e.parameter.phone;
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("Users");
  const values = sheet.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {
    if (values[i][2] == phone) {
      sheet.getRange(i+1, 7).setValue(""); // Clear DeviceID in Column G
      break;
    }
  }
  return jsonResponse({status: "success"});
}

// ========== 4. PDF + VIDEO FUNCTIONS ==========
function getPDFs(e){
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("PDFs") || ss.insertSheet("PDFs");
  if(sheet.getLastRow() < 2) return jsonResponse({pdfs: []});
  const data = sheet.getDataRange().getValues();
  const pdfs = [];
  for (let i = 1; i < data.length; i++) {
    pdfs.push({Title: data[i][0], Level: data[i][2], Subject: "", DriveLink: data[i][1]});
  }
  return jsonResponse({pdfs: pdfs});
}

function addPDF(data){
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("PDFs") || ss.insertSheet("PDFs");
  sheet.appendRow([data.Title, data.Link, data.Class, new Date()]);
  return jsonResponse({status: "success"});
}

function deletePDF(data){
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("PDFs");
  sheet.deleteRow(data.Row);
  return jsonResponse({status: "success"});
}

function getVideos(e){
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("Videos") || ss.insertSheet("Videos");
  if(sheet.getLastRow() < 2) return jsonResponse({videos: []});
  const data = sheet.getDataRange().getValues();
  const videos = [];
  for (let i = 1; i < data.length; i++) {
    videos.push({Title: data[i][0], Level: data[i][2], Subject: "", VideoLink: data[i][1]});
  }
  return jsonResponse({videos: videos});
}

function addVideo(data){
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("Videos") || ss.insertSheet("Videos");
  sheet.appendRow([data.Title, data.Link, data.Class, new Date()]);
  return jsonResponse({status: "success"});
}

function deleteVideo(data){
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("Videos");
  sheet.deleteRow(data.Row);
  return jsonResponse({status: "success"});
}

// ========== HELPER ==========
function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
  .setMimeType(ContentService.MimeType.JSON);
}