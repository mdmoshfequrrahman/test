// Code.gs

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Akij Light Engineering - Production App')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 1. Get User Profile & Four-Tier Access Logic
function getUserState() {
  var userEmail = Session.getActiveUser().getEmail();
  
  // NOTE: To get Name & Photo, enable "Admin Directory API" in Apps Script Services
  var userName = userEmail.split('@')[0].toUpperCase(); // Fallback name
  var userPhoto = "https://ui-avatars.com/api/?name=" + userName + "&background=00E676&color=121212"; // Fallback Avatar
  
  try {
    // If Workspace API is enabled, this pulls official company directory info
    var user = AdminDirectory.Users.get(userEmail);
    userName = user.name.fullName;
    userPhoto = user.thumbnailPhotoUrl || userPhoto;
  } catch(e) { /* Ignore if API not enabled yet */ }

  var access = { email: userEmail, name: userName, photo: userPhoto, role: "Viewer", lines: "ALL" };
  
  // Tier 1: Hardcoded Super Admin
  var superAdminEmail = "head.plant@akijlighengineering.com"; // Change to your email
  if (userEmail === superAdminEmail) {
    access.role = "Super Admin";
    access.lines = "ALL";
    return compileAppState(access);
  }
  
  // Tiers 2-4: Check Database for Admins, Editors (Data Entry), and Viewers
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Config_Users");
  var data = sheet.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toLowerCase() === userEmail.toLowerCase()) {
      access.role = data[i][1]; // "Admin", "Data Entry", "Viewer"
      access.lines = data[i][2]; // e.g., "L-01, L-02"
      break;
    }
  }
  return compileAppState(access);
}

// 2. Package everything the Frontend needs to load instantly
function compileAppState(access) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Fetch dropdown data for the Data Entry form
  var linesData = ss.getSheetByName("Config_Lines").getDataRange().getValues();
  var productsData = ss.getSheetByName("Config_Products").getDataRange().getValues();
  
  return {
    user: access,
    linesConfig: linesData,
    productsConfig: productsData
  };
}

// 3. API Entry Point: Submit Production Data (For Supervisors)
function submitProductionData(lineId, productId, hourSlot, actualOutput) {
  var access = getUserState().user;
  
  // Security Check: Only Data Entry, Admins, or Super Admins can submit
  if (access.role === "Viewer") throw new Error("Viewers cannot submit data.");
  
  // Double check if this user is assigned to this specific line (unless Admin)
  if (access.role === "Data Entry" && access.lines.indexOf(lineId) === -1) {
      throw new Error("You do not have permission to submit data for this line.");
  }
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Production_Log");
  var timestamp = new Date();
  var dateStr = Utilities.formatDate(timestamp, "GMT+6", "yyyy-MM-dd");
  
  sheet.appendRow([timestamp, dateStr, lineId, productId, hourSlot, actualOutput, access.email]);
  return "Data saved successfully for " + hourSlot;
}
