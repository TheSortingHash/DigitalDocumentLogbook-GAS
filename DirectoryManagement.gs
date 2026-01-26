/**
 * HANDLES DIRECTORY CSV UPLOAD
 * 1. Parses CSV data handling quoted strings (e.g., "Doe, John").
 * 2. Clears the existing Directory sheet (preserving headers).
 * 3. Bulk writes the new data.
 */
function processDirectoryUpload(csvContent) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000); // Wait up to 30 seconds
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Directory');
    if (!sheet) throw new Error("Critical: 'Directory' sheet not found.");

    // Utilities.parseCsv handles quotes/commas automatically
    const parsedData = Utilities.parseCsv(csvContent);
    
    if (parsedData.length === 0) return { success: false, error: "File is empty." };

    // Remove Header row if it exists in the CSV (Check if A1 is 'Name')
    let dataToWrite = parsedData;
    if (parsedData[0][0].toLowerCase().trim() === 'name') {
      dataToWrite = parsedData.slice(1);
    }

    if (dataToWrite.length === 0) return { success: false, error: "No data rows found in CSV." };

    // Clear existing data (Row 2 onwards, Columns A & B)
    // We leave Row 1 (Headers) intact.
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, 2).clearContent();
    }

    // Write new data in bulk
    sheet.getRange(2, 1, dataToWrite.length, dataToWrite[0].length).setValues(dataToWrite);

    return { 
      success: true, 
      count: dataToWrite.length, 
      preview: dataToWrite.slice(0, 3) // Return top 3 for preview
    };

  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    lock.releaseLock();
  }
}