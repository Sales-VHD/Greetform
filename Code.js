const DATA_ENTRY_SHEET_NAME = "Sheet1";
const TIME_STAMP_COLUMN_NAME = "Timestamp";
const FOLDER_ID = "";
const FILE_LINK_COLUMN_NAME = "FileLink";
const UPLOADED_FILE_NAME_COLUMN = "UploadedFileName";
const FORM_FILE_INPUT_NAME = "theFile";

function doPost(e) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
      DATA_ENTRY_SHEET_NAME
    );
    if (!sheet) {
      throw new Error(`Sheet '${DATA_ENTRY_SHEET_NAME}' not found`);
    }

    const formData = e.postData.contents ? JSON.parse(e.postData.contents) : {};

    // Handle file upload if present
    let fileInfo = null;
    if (formData.fileData) {
      fileInfo = saveFile(formData.fileData);
      delete formData.fileData; // Remove file data from form data
    }

    // Prepare data for sheet
    const rowData = {
      ...formData,
      [TIME_STAMP_COLUMN_NAME]: new Date().toISOString(),
    };

    if (fileInfo) {
      rowData[FILE_LINK_COLUMN_NAME] = fileInfo.url;
      rowData[UPLOADED_FILE_NAME_COLUMN] = fileInfo.name;
    }

    appendToGoogleSheet(rowData, sheet);

    return ContentService.createTextOutput(
      JSON.stringify({
        status: "success",
        message: "Data submitted successfully",
      })
    ).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    console.error(error);
    return ContentService.createTextOutput(
      JSON.stringify({
        status: "error",
        message: error.toString(),
      })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Saves a file to Google Drive
 */
function saveFile(fileData) {
  try {
    const blob = Utilities.newBlob(
      Utilities.base64Decode(fileData.data),
      fileData.mimeType,
      fileData.fileName
    );
    const folder = DriveApp.getFolderById(FOLDER_ID);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return {
      url: `https://drive.google.com/uc?export=view&id=${file.getId()}`,
      name: fileData.fileName,
    };
  } catch (error) {
    console.error("File upload error:", error);
    throw new Error("Failed to upload file: " + error.toString());
  }
}

/**
 * Appends data to the Google Sheet
 */
function appendToGoogleSheet(data, sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  // If sheet is empty, create headers
  if (headers.length === 0 || headers[0] === "") {
    const newHeaders = Object.keys(data);
    sheet.getRange(1, 1, 1, newHeaders.length).setValues([newHeaders]);
    headers = newHeaders;
  }

  // Map data to header columns
  const rowData = headers.map((header) => data[header] || "");
  sheet.appendRow(rowData);
}
