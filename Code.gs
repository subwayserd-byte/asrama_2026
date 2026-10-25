var IMAGE_FOLDER_IDS = {
  aduan: "1LPYl-WuS_7jB0lymTjkFal6L39hFVZgU",
  kendiri: "1rbdUKTGCSYRBjfOPowr2czcnc-vNGc1D"
};
var IMAGE_RETENTION_DAYS = 7;

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ticketId = String(data.id || "").trim();
    var action = String(data.action || "");

    if (action === "login") {
      return handleAdminLogin_(data);
    }

    if (action === "logout") {
      if (!isValidAdminSession_(data.token)) return jsonResponse_({ result: "unauthorized" });
      CacheService.getScriptCache().remove("admin_session_" + data.token);
      return jsonResponse_({ result: "success" });
    }

    if (action === "delete" || action === "deleteAll" || action === "update") {
      if (!isValidAdminSession_(data.token)) return jsonResponse_({ result: "unauthorized" });
    }

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var rows = sheet.getDataRange().getValues();

    if (action === "delete") {
      var matchingRows = [];
      for (var matchIndex = 1; matchIndex < rows.length; matchIndex++) {
        if (rows[matchIndex][0] &&
            String(rows[matchIndex][0]).trim().toLowerCase() === ticketId.toLowerCase()) {
          matchingRows.push({
            rowNumber: matchIndex + 1,
            imageUrl: String(rows[matchIndex][9] || ""),
            reportType: String(rows[matchIndex][10] || "aduan")
          });
        }
      }

      var trashedImageCount = 0;
      matchingRows.forEach(function(record) {
        if (record.imageUrl && trashStoredImage_(record.imageUrl, record.reportType)) {
          trashedImageCount++;
        }
      });

      var deletedCount = 0;
      for (var removeIndex = matchingRows.length - 1; removeIndex >= 0; removeIndex--) {
        sheet.deleteRow(matchingRows[removeIndex].rowNumber);
        deletedCount++;
      }
      return ContentService.createTextOutput(JSON.stringify({
        result: "success",
        deleted: deletedCount,
        trashedImages: trashedImageCount
      })).setMimeType(ContentService.MimeType.JSON);
    }

    if (action === "deleteAll") {
      var deleteIds = Array.isArray(data.ids) ? data.ids.map(function(id) {
        return String(id || "").trim().toLowerCase();
      }).filter(function(id) {
        return !!id;
      }) : [];

      if (!deleteIds.length) {
        return ContentService.createTextOutput(JSON.stringify({
          result: "success",
          deleted: 0,
          trashedImages: 0
        })).setMimeType(ContentService.MimeType.JSON);
      }

      var bulkMatchingRows = [];
      for (var bulkIndex = 1; bulkIndex < rows.length; bulkIndex++) {
        var rowId = rows[bulkIndex][0];
        if (!rowId) continue;
        var normalizedRowId = String(rowId).trim().toLowerCase();
        if (deleteIds.indexOf(normalizedRowId) !== -1) {
          bulkMatchingRows.push({
            rowNumber: bulkIndex + 1,
            imageUrl: String(rows[bulkIndex][9] || ""),
            reportType: String(rows[bulkIndex][10] || "aduan")
          });
        }
      }

      var bulkTrashedImageCount = 0;
      bulkMatchingRows.forEach(function(record) {
        if (record.imageUrl && trashStoredImage_(record.imageUrl, record.reportType)) {
          bulkTrashedImageCount++;
        }
      });

      var bulkDeletedCount = 0;
      for (var bulkRemoveIndex = bulkMatchingRows.length - 1; bulkRemoveIndex >= 0; bulkRemoveIndex--) {
        sheet.deleteRow(bulkMatchingRows[bulkRemoveIndex].rowNumber);
        bulkDeletedCount++;
      }

      return ContentService.createTextOutput(JSON.stringify({
        result: "success",
        deleted: bulkDeletedCount,
        trashedImages: bulkTrashedImageCount
      })).setMimeType(ContentService.MimeType.JSON);
    }

    var foundRowIndex = -1;
    for (var rowIndex = 1; rowIndex < rows.length; rowIndex++) {
      if (rows[rowIndex][0] && String(rows[rowIndex][0]).trim().toLowerCase() === ticketId.toLowerCase()) {
        foundRowIndex = rowIndex + 1;
        break;
      }
    }

    if (foundRowIndex > 0 && !isValidAdminSession_(data.token)) {
      return jsonResponse_({ result: "unauthorized" });
    }

    if (data.reportType === "kendiri" && foundRowIndex < 0) {
      var requestedRoom = String(data.blok || data.Wing || "").trim().toLowerCase();
      var requestedWeek = getIsoWeekKey_(data.tarikh);
      for (var checkIndex = 1; checkIndex < rows.length; checkIndex++) {
        var existingType = String(rows[checkIndex][10] || "aduan");
        var existingRoom = String(rows[checkIndex][3] || "").trim().toLowerCase();
        if (existingType === "kendiri" &&
            existingRoom === requestedRoom &&
            getIsoWeekKey_(rows[checkIndex][8]) === requestedWeek) {
          return ContentService.createTextOutput(JSON.stringify({
            result: "duplicate",
            message: "Laporan bilik ini telah dihantar untuk minggu semasa."
          })).setMimeType(ContentService.MimeType.JSON);
        }
      }
    }

    var imageUrl = foundRowIndex > 0 ? (rows[foundRowIndex - 1][9] || "") : "";
    var imageData = String(data.gambarBase64 || data.gambar || "");
    var imageIsMissing = !imageUrl || imageUrl === "Tiada Gambar" || String(imageUrl).indexOf("Gagal Simpan Gambar") === 0;

    if (imageData.indexOf("data:image/") === 0 && (foundRowIndex < 0 || imageIsMissing)) {
      try {
        imageUrl = saveComplaintImage_(data, imageData);
      } catch (imageError) {
        imageUrl = "Gagal Simpan Gambar: " + (imageError.message || String(imageError));
      }
    } else if (!imageUrl) {
      imageUrl = "Tiada Gambar";
    }

    if (foundRowIndex > 0) {
      sheet.getRange(foundRowIndex, 8).setValue(data.status || rows[foundRowIndex - 1][7]);
      sheet.getRange(foundRowIndex, 9).setValue(data.tarikh || rows[foundRowIndex - 1][8] || new Date());
      if (imageData.indexOf("data:image/") === 0 && imageUrl !== "Tiada Gambar") {
        sheet.getRange(foundRowIndex, 10).setValue(imageUrl);
      }
    } else {
      sheet.appendRow([
        data.id,
        data.nama,
        data.matrik,
        data.blok || data.Wing || "",
        data.telefon,
        data.kategori,
        data.keterangan,
        data.status || "Pending",
        data.tarikh || new Date(),
        imageUrl,
        data.reportType || "aduan",
        data.condition || ""
      ]);
    }

    return ContentService.createTextOutput(JSON.stringify({ result: "success" }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      result: "error",
      message: String(err)
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  var params = e && e.parameter ? e.parameter : {};
  if (params.action === "track") {
    return getPublicTicketStatus_(params.id);
  }
  if (params.action === "checkRoomReport") {
    return checkPublicRoomReport_(params.room, params.week);
  }

  if (params.action !== "admin" || !isValidAdminSession_(params.token)) {
    return jsonResponse_({ result: "unauthorized" });
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var rows = sheet.getDataRange().getValues();
  var data = [];

  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0]) {
      data.push({
        id: rows[i][0],
        nama: rows[i][1],
        matrik: rows[i][2],
        blok: rows[i][3],
        telefon: rows[i][4],
        kategori: rows[i][5],
        keterangan: rows[i][6],
        status: rows[i][7],
        tarikh: rows[i][8],
        gambarBase64: rows[i][9] || "",
        reportType: rows[i][10] || "aduan",
        condition: rows[i][11] || ""
      });
    }
  }

  return jsonResponse_(data);
}

function jsonResponse_(value) {
  return ContentService.createTextOutput(JSON.stringify(value))
    .setMimeType(ContentService.MimeType.JSON);
}

function handleAdminLogin_(data) {
  var cache = CacheService.getScriptCache();
  var username = String(data.username || "").trim().toLowerCase();
  var password = String(data.password || "");
  var token = String(data.token || "");
  var callerKey = getLoginRateLimitKey_();
  var failedAttempts = Number(cache.get(callerKey) || 0);

  if (failedAttempts >= 5) {
    return jsonResponse_({ result: "error", message: "Terlalu banyak cubaan. Cuba semula dalam 15 minit." });
  }

  var properties = PropertiesService.getScriptProperties();
  var configuredUsername = String(properties.getProperty("ADMIN_USERNAME") || "").trim().toLowerCase();
  var salt = properties.getProperty("ADMIN_PASSWORD_SALT") || "";
  var expectedHash = properties.getProperty("ADMIN_PASSWORD_HASH") || "";
  var version = properties.getProperty("ADMIN_CREDENTIAL_VERSION") || "";

  if (!configuredUsername || !salt || !expectedHash || !version) {
    return jsonResponse_({ result: "not_configured", message: "Admin belum dikonfigurasi di Apps Script." });
  }

  var suppliedHash = hashAdminPassword_(salt, password);
  var credentialsMatch = username === configuredUsername &&
    constantTimeEquals_(suppliedHash, expectedHash);
  var tokenIsValid = /^[a-f0-9]{64}$/i.test(token);

  if (!credentialsMatch || !tokenIsValid) {
    cache.put(callerKey, String(failedAttempts + 1), 900);
    return jsonResponse_({ result: "error", message: "Username atau kata laluan tidak sah." });
  }

  cache.remove(callerKey);
  cache.put("admin_session_" + token, version, 21600);
  return jsonResponse_({ result: "success" });
}

function isValidAdminSession_(token) {
  var sessionToken = String(token || "");
  if (!/^[a-f0-9]{64}$/i.test(sessionToken)) return false;

  var sessionVersion = CacheService.getScriptCache().get("admin_session_" + sessionToken);
  var currentVersion = PropertiesService.getScriptProperties().getProperty("ADMIN_CREDENTIAL_VERSION");
  return !!sessionVersion && !!currentVersion && sessionVersion === currentVersion;
}

function getPublicTicketStatus_(ticketId) {
  var requestedId = String(ticketId || "").trim().toLowerCase();
  if (!/^#?tvet-\d{4}$/.test(requestedId)) {
    return jsonResponse_({ result: "not_found" });
  }

  function checkPublicRoomReport_(room, weekKey) {
    var requestedRoom = String(room || "").trim().toLowerCase();
    var requestedWeek = String(weekKey || "");
    if (!requestedRoom || requestedRoom.length > 100 || !/^\d{4}-\d{2}$/.test(requestedWeek)) {
      return jsonResponse_({ result: "error", message: "Maklumat bilik atau minggu tidak sah." });
    }

    var rows = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      var existingType = String(rows[i][10] || "aduan");
      var existingRoom = String(rows[i][3] || "").trim().toLowerCase();
      if (existingType === "kendiri" &&
          existingRoom === requestedRoom &&
          getIsoWeekKey_(rows[i][8]) === requestedWeek) {
        return jsonResponse_({ result: "success", exists: true });
      }
    }

    return jsonResponse_({ result: "success", exists: false });
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0] || "").trim().toLowerCase() === requestedId) {
      return jsonResponse_({
        result: "success",
        id: String(rows[i][0]),
        status: String(rows[i][7] || "Pending")
      });
    }
  }

  return jsonResponse_({ result: "not_found" });
}

function getLoginRateLimitKey_() {
  var userKey = "";
  try {
    userKey = Session.getTemporaryActiveUserKey() || "";
  } catch (error) {
    userKey = "";
  }
  return "admin_login_failures_" + (userKey || "anonymous");
}

function hashAdminPassword_(salt, password) {
  var digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    salt + ":" + password,
    Utilities.Charset.UTF_8
  );
  return Utilities.base64EncodeWebSafe(digest);
}

function constantTimeEquals_(left, right) {
  var a = String(left || "");
  var b = String(right || "");
  var difference = a.length ^ b.length;
  var length = Math.max(a.length, b.length);
  for (var i = 0; i < length; i++) {
    difference |= (a.charCodeAt(i) || 0) ^ (b.charCodeAt(i) || 0);
  }
  return difference === 0;
}

function configureAdminCredentials() {
  var ui = SpreadsheetApp.getUi();
  var usernameResponse = ui.prompt(
    "Tetapkan admin baharu",
    "Masukkan username admin baharu:",
    ui.ButtonSet.OK_CANCEL
  );
  if (usernameResponse.getSelectedButton() !== ui.Button.OK) return;

  var username = usernameResponse.getResponseText().trim();
  if (!username || username.length > 80) {
    throw new Error("Username mesti mengandungi 1 hingga 80 aksara.");
  }

  var passwordResponse = ui.prompt(
    "Tetapkan kata laluan admin",
    "Masukkan kata laluan baharu (sekurang-kurangnya 10 aksara):",
    ui.ButtonSet.OK_CANCEL
  );
  if (passwordResponse.getSelectedButton() !== ui.Button.OK) return;

  var password = passwordResponse.getResponseText();
  if (password.length < 10) {
    throw new Error("Kata laluan mesti sekurang-kurangnya 10 aksara.");
  }

  var properties = PropertiesService.getScriptProperties();
  var salt = Utilities.getUuid() + Utilities.getUuid();
  properties.setProperties({
    ADMIN_USERNAME: username,
    ADMIN_PASSWORD_SALT: salt,
    ADMIN_PASSWORD_HASH: hashAdminPassword_(salt, password),
    ADMIN_CREDENTIAL_VERSION: Utilities.getUuid()
  });

  ui.alert("Kelayakan admin dikemas kini. Sesi admin lama telah dibatalkan.");
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Admin Portal")
    .addItem("Tetapkan kelayakan admin", "configureAdminCredentials")
    .addToUi();
}

function getIsoWeekKey_(value) {
  var match = String(value || "").match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!match) return "";
  var date = new Date(Date.UTC(Number(match[1]), Number(match[2]) - 1, Number(match[3])));
  var day = date.getUTCDay() || 7;
  date.setUTCDate(date.getUTCDate() + 4 - day);
  var yearStart = new Date(Date.UTC(date.getUTCFullYear(), 0, 1));
  var week = Math.ceil((((date - yearStart) / 86400000) + 1) / 7);
  return date.getUTCFullYear() + "-" + ("0" + week).slice(-2);
}

function saveComplaintImage_(data, imageData) {
  var match = imageData.match(/^data:(image\/[a-zA-Z0-9.+-]+);base64,([\s\S]+)$/);
  if (!match) throw new Error("Format data gambar tidak sah.");

  var reportType = data.reportType === "kendiri" ? "kendiri" : "aduan";
  var folderId = IMAGE_FOLDER_IDS[reportType];
  var folder = DriveApp.getFolderById(folderId);
  var location = String(data.blok || data.Wing || "Bilik")
    .trim()
    .replace(/[^A-Za-z0-9._-]+/g, "_")
    .replace(/\s+/g, "_")
    .substring(0, 70);
  var id = String(data.id || "Tanpa_ID")
    .replace(/^#/, "")
    .replace(/[^A-Za-z0-9_-]/g, "-");
  var date = String(data.tarikh || new Date().toISOString().slice(0, 10)).slice(0, 10);
  var prefix = reportType === "kendiri" ? "Kendiri" : "Aduan";
  var extension = match[1] === "image/png" ? "png" : "jpg";
  var mimeType = match[1];
  var fileName = prefix + "_" + location + "_" + id + "_" + date + "." + extension;
  var blob = Utilities.newBlob(Utilities.base64Decode(match[2]), mimeType, fileName);
  var file = folder.createFile(blob);
  file.setDescription("Akan dipadam secara automatik selepas " + IMAGE_RETENTION_DAYS + " hari.");
  return file.getUrl();
}

function trashStoredImage_(imageUrl, reportType) {
  var url = String(imageUrl || "");
  var match = url.match(/\/d\/([A-Za-z0-9_-]+)/) || url.match(/[?&]id=([A-Za-z0-9_-]+)/);
  if (!match || url.indexOf("https://drive.google.com/") !== 0) return false;

  var expectedType = reportType === "kendiri" ? "kendiri" : "aduan";
  var expectedPrefix = expectedType === "kendiri" ? "Kendiri_" : "Aduan_";
  var file = DriveApp.getFileById(match[1]);
  if (file.isTrashed() || file.getName().indexOf(expectedPrefix) !== 0) return false;

  var parents = file.getParents();
  var isInExpectedFolder = false;
  while (parents.hasNext()) {
    if (parents.next().getId() === IMAGE_FOLDER_IDS[expectedType]) {
      isInExpectedFolder = true;
      break;
    }
  }

  if (!isInExpectedFolder) return false;
  file.setTrashed(true);
  return true;
}

function cleanupOldComplaintImages() {
  var cutoff = Date.now() - IMAGE_RETENTION_DAYS * 24 * 60 * 60 * 1000;
  Object.keys(IMAGE_FOLDER_IDS).forEach(function(type) {
    var files = DriveApp.getFolderById(IMAGE_FOLDER_IDS[type]).getFiles();
    while (files.hasNext()) {
      var file = files.next();
      if (/^(Aduan|Kendiri)_/.test(file.getName()) && file.getDateCreated().getTime() < cutoff) {
        file.setTrashed(true);
      }
    }
  });
}

function installImageCleanupTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === "cleanupOldComplaintImages") {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  ScriptApp.newTrigger("cleanupOldComplaintImages")
    .timeBased()
    .everyHours(1)
    .create();
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Admin Portal")
    .addItem("Tetapkan kelayakan admin", "configureAdminCredentials")
    .addToUi();
}
