

/**
 * Sends email for Internal Handover.
 * Updated to handle BATCH documents (Array) and safer image attachment.
 */
function sendRoutedEmail(recipientData, docsList, signatureBlob, logID, timestamp) {
  try {
    const subject = `[Finance] Documents Routed - Ref: ${logID}`;
    const formattedTimestamp = timestamp.toLocaleString('en-US', { timeZone: 'Asia/Manila' });

    // 1. Build Table Rows
    const rows = docsList.map(d => 
      `<tr>
         <td style="padding:12px;border-bottom:1px solid #eee;">${d.Title}</td>
         <td style="padding:12px;border-bottom:1px solid #eee;">${d.Type}</td>
       </tr>`
    ).join('');

    // 2. Handle Signature Safety
    let signatureHtml = '';
    let inlineImagesObj = {};

    if (signatureBlob) {
      signatureHtml = `<img src="cid:signatureImage" style="height:80px;">`;
      inlineImagesObj['signatureImage'] = signatureBlob;
    } else {
      signatureHtml = `<p style="color:#888; font-style:italic;">(Digital Signature Pending)</p>`;
    }

    const html = `
      <div style="background-color: #f4f6f8; padding: 40px 0; font-family: Arial, sans-serif;">
        <div style="max-width: 650px; margin: 0 auto; background-color: #ffffff; border-radius: 8px; overflow: hidden;">
          <div style="background-color: #1C2790; padding: 30px; text-align: center;">
            <img src="https://i.imgur.com/jaEbfAR.png" width="400">
          </div>
          <div style="padding: 40px; border-top: 6px solid #1C2790;">
            <h1 style="color: #1C2790; text-align: center;">DOCUMENTS ROUTED</h1>
            <p>Dear <b>${recipientData.name}</b>,</p>
            <p>This email confirms that the following documents have been routed to your custody:</p>
            
            <div style="background-color:#f8f9fa; padding:10px; margin-bottom:15px; border-left:4px solid #1C2790;">
              <strong>Log ID:</strong> ${logID}<br>
              <strong>Remarks:</strong> ${docsList[0].Remarks || "N/A"}
            </div>

            <table style="width:100%;border-collapse:collapse;margin-top:10px;">
              <thead>
                <tr style="background-color:#f2f4f8;">
                  <th style="padding:10px; text-align:left;">Title</th>
                  <th style="padding:10px; text-align:left;">Type</th>
                </tr>
              </thead>
              <tbody>${rows}</tbody>
            </table>

            <div style="margin-top:30px;border:1px solid #eee;padding:20px;text-align:center;">
              <p style="font-weight:bold; color:#555;">RECEIVED BY</p>
              <div style="display:inline-block;padding:5px;border:1px dashed #ccc;">
                ${signatureHtml}
              </div>
              <p style="margin-top:10px; font-weight:bold;">${recipientData.name}</p>
              <p style="font-size:12px; color:#777;">Date: ${formattedTimestamp}</p>
            </div>
          </div>
          <div style="background-color: #eeeeee; padding: 20px; text-align: center; color: #888;">
            &copy; Finance Department - Office of the Dept Manager.
          </div>
        </div>
      </div>`;

    MailApp.sendEmail({
      to: recipientData.email,
      subject: subject,
      htmlBody: html,
      name: 'Finance - Internal Routing',
      inlineImages: inlineImagesObj // Use the safe object
    });
    
    console.log(`Email sent successfully to ${recipientData.email}`);

  } catch (e) {
    console.error("sendRoutedEmail Failed: " + e.toString());
    // Do not throw error so the main process doesn't crash
  }
}




/**
 * Generates unique ID based on InternalTransactions sheet.
 * Format: LOG-YYYY-MM-000
 */
function generateLogID() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("InternalTransactions"); // CHANGED SHEET NAME
  if (!sheet) return "LOG-ERROR";

  const data = sheet.getDataRange().getValues();
  const today = new Date();
  const year = today.getFullYear();
  const month = String(today.getMonth() + 1).padStart(2, '0');
  const prefix = `LOG-${year}-${month}`;
  
  let maxSeq = 0;
  
  // Skip header (i=1)
  for (let i = 1; i < data.length; i++) {
    const id = String(data[i][0]);
    if (id.startsWith(prefix)) {
      const parts = id.split('-');
      const seq = parseInt(parts[3], 10);
      if (!isNaN(seq) && seq > maxSeq) {
        maxSeq = seq;
      }
    }
  }

  const newSeq = String(maxSeq + 1).padStart(3, '0');
  return `${prefix}-${newSeq}`;
}



/**
 * Checks whether any of the supplied full titles already exist in InternalDocuments.
 * Used to warn when routing the same external document into the internal logbook twice.
 */
function checkExternalRefDuplicates(fullTitles) {
  try {
    if (!fullTitles || !fullTitles.length) return [];
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const docSheet = ss.getSheetByName("InternalDocuments");
    if (!docSheet || docSheet.getLastRow() < 2) return [];

    const existing = docSheet.getRange(2, 2, docSheet.getLastRow() - 1, 1)
                             .getValues().flat()
                             .map(t => String(t).trim().toLowerCase());

    const seen = new Set(existing);
    return fullTitles.filter(t => seen.has(String(t).trim().toLowerCase()));
  } catch (e) {
    console.error("checkExternalRefDuplicates failed: " + e.toString());
    return [];
  }
}


/* --- INTERNAL LOGBOOK: SERVER SIDE --- */

// 1. GET DATA FOR INTERNAL UPDATE PAGE (Mirrors getTransactionDetails)
function getInternalLogDetails(logId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const txSheet = ss.getSheetByName("InternalTransactions");
  const docSheet = ss.getSheetByName("InternalDocuments");
  
  if (!txSheet || !docSheet) return null;

  // Get Header Info
  const txData = txSheet.getDataRange().getValues();
  let logHeader = null;
  for (let i = 1; i < txData.length; i++) {
    if (String(txData[i][0]) === String(logId)) {
      logHeader = {
        LogID: txData[i][0],
        Timestamp: txData[i][1],
        Encoder: txData[i][2]
      };
      break;
    }
  }
  
  if (!logHeader) return null;

  // Get Document Details
  const docData = docSheet.getDataRange().getValues();
  const documents = [];
  
  for (let i = 1; i < docData.length; i++) {
    if (String(docData[i][0]) === String(logId)) {
      documents.push({
        rowNumber: i + 1, // Store Excel row number for updates
        Title: docData[i][1],
        Type: docData[i][2],
        Status: docData[i][3],
        CurrentCustody: docData[i][4],
        History: docData[i][5]
      });
    }
  }

  return { header: logHeader, documents: documents };
}

// 2. PROCESS INITIAL LOGGING (Multi-Doc, Finance Custody)
function processInternalLog(formData) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000); 

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const txSheet = ss.getSheetByName("InternalTransactions");
    const docSheet = ss.getSheetByName("InternalDocuments");
    
    const logID = generateLogID(); 
    const timestamp = new Date();
    
    // A. Save Header
    txSheet.appendRow([
      logID,
      timestamp,
      formData.encoder || "Finance Staff",
      formData.documents.length
    ]);

    // B. Handle "Received From" (Optional Email)
    let historyNote = `[${timestamp.toLocaleString()}] Logged by Finance`;
    if (formData.receivedFrom) {
       historyNote += ` [Received from: ${formData.receivedFrom}]`;
       
       // Lookup and Email
       const directory = getDirectory();
       const targetName = String(formData.receivedFrom).trim().toLowerCase();
       const recipientObj = directory.find(d => String(d.name).trim().toLowerCase() === targetName);
       
       if (recipientObj && recipientObj.email) {
          sendFinanceReceivedEmail(
             { name: formData.receivedFrom, email: recipientObj.email },
             formData.documents,
             logID,
             false // isReturn = false
          );
       }
    }

    // C. Save Documents
    formData.documents.forEach(doc => {
      docSheet.appendRow([
        logID,
        doc.docTitle,
        doc.docType,
        "Finance Custody", 
        "Finance",         
        historyNote, 
        ""                 
      ]);
    });

    return { 
      success: true, 
      logID: logID, 
      timestamp: timestamp.toLocaleString(),
      appUrl: ScriptApp.getService().getUrl(),
      documents: formData.documents 
    };

  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

/* --- FIXED HANDOVER FUNCTION (Passes Array Correctly) --- */
function processInternalHandover(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const docSheet = ss.getSheetByName("InternalDocuments");
    
    // Determine if this is a "Return to Finance" action
    const isReturn = (data.recipientName === "Finance");
    
    let fileUrl = "";
    // Only process signature if NOT returning to Finance
    if (!isReturn && data.signature) {
        const folder = DriveApp.getFoldersByName('DMS Signatures').hasNext() ? DriveApp.getFoldersByName('DMS Signatures').next() : DriveApp.createFolder('DMS Signatures');
        const imageData = data.signature.split(',')[1];
        const blob = Utilities.newBlob(Utilities.base64Decode(imageData), 'image/png', `${data.logID}-Handover.png`);
        const file = folder.createFile(blob);
        fileUrl = file.getUrl();
    }

    const timestamp = new Date();
    const affectedDocs = [];
    let previousOwnerName = ""; // To notify them if returning

    // Update Rows
    data.docRows.forEach(rowIndex => {
      const row = parseInt(rowIndex);
      
      // Capture previous owner before updating (for return notification)
      if (isReturn && !previousOwnerName) {
         previousOwnerName = docSheet.getRange(row, 5).getValue();
      }

      // Logic: If Return -> Status "Finance Custody", Custody "Finance"
      // If Handover -> Status "Routed", Custody [Name]
      const newStatus = isReturn ? "Finance Custody" : "Routed";
      const newCustody = isReturn ? "Finance" : data.recipientName;

      docSheet.getRange(row, 4).setValue(newStatus);
      docSheet.getRange(row, 5).setValue(newCustody);
      
      const currentHist = docSheet.getRange(row, 6).getValue();
      const actionText = isReturn ? "Returned to Finance" : `Passed to ${data.recipientName}`;
      const newHist = `${currentHist}\n[${timestamp.toLocaleString()}] ${actionText} (${data.remarks})`;
      
      docSheet.getRange(row, 6).setValue(newHist);
      if(fileUrl) docSheet.getRange(row, 7).setValue(fileUrl);

      affectedDocs.push({
        Title: docSheet.getRange(row, 2).getValue(),
        Type: docSheet.getRange(row, 3).getValue(),
        Remarks: data.remarks
      });
    });

    // --- NOTIFICATION LOGIC ---
    const directory = getDirectory();
    
    if (isReturn) {
       // Notify the person who returned it (Previous Owner)
       const targetName = String(previousOwnerName).trim().toLowerCase();
       const prevOwnerObj = directory.find(d => String(d.name).trim().toLowerCase() === targetName);
       
       if (prevOwnerObj && prevOwnerObj.email) {
          sendFinanceReceivedEmail(
             { name: previousOwnerName, email: prevOwnerObj.email },
             affectedDocs,
             data.logID,
             true // isReturn = true
          );
       }
    } else {
       // Standard Handover Notification (with Signature)
       const targetName = String(data.recipientName).trim().toLowerCase();
       const recipientObj = directory.find(d => String(d.name).trim().toLowerCase() === targetName);
       
       if (recipientObj && recipientObj.email) {
          // Re-fetch blob for email if it exists
          let blob = null;
          if (data.signature) {
             const imageData = data.signature.split(',')[1];
             blob = Utilities.newBlob(Utilities.base64Decode(imageData), 'image/png', 'sig.png');
          }

          sendRoutedEmail(
            { name: data.recipientName, email: recipientObj.email },
            affectedDocs, 
            blob,
            data.logID,
            timestamp
          );
       }
    }

    return { success: true };

  } catch (e) {
    console.error("Handover Error: " + e.toString());
    return { success: false, error: e.toString() };
  }
}

/**
 * Sends acknowledgment when Finance receives documents 
 * (either Initial Log from External OR Returned from Internal Staff).
 */
function sendFinanceReceivedEmail(recipientData, docsList, logID, isReturn) {
  const subject = `[Finance] Document Receipt Acknowledgement - Ref: ${logID}`;
  const timestamp = new Date().toLocaleString('en-US', { timeZone: 'Asia/Manila' });
  
  // Dynamic Header Text
  const headerText = isReturn ? "RETURNED TO FINANCE" : "RECEIPT ACKNOWLEDGEMENT";
  const introText = isReturn 
    ? "This email confirms that the Finance Department has received the following documents back from your custody:"
    : "This email serves as proof that the Finance Department has received the following documents from you:";

  const rows = docsList.map(d => 
    `<tr>
       <td style="padding:12px;border-bottom:1px solid #eee;">${d.Title || d.docTitle}</td>
       <td style="padding:12px;border-bottom:1px solid #eee;">${d.Type || d.docType}</td>
     </tr>`
  ).join('');

  const html = `
    <div style="background-color: #f4f6f8; padding: 40px 0; font-family: Arial, sans-serif;">
      <div style="max-width: 650px; margin: 0 auto; background-color: #ffffff; border-radius: 8px; overflow: hidden;">
        <div style="background-color: #1C2790; padding: 30px; text-align: center;">
          <img src="https://i.imgur.com/jaEbfAR.png" width="400">
        </div>
        <div style="padding: 40px; border-top: 6px solid #1C2790;">
          <h1 style="color: #1C2790; text-align: center;">${headerText}</h1>
          <p>Dear <b>${recipientData.name}</b>,</p>
          <p>${introText}</p>
          
          <div style="background-color:#f8f9fa; padding:10px; margin-bottom:15px; border-left:4px solid #1C2790;">
            <strong>Log ID:</strong> ${logID}<br>
            <strong>Date Received:</strong> ${timestamp}
          </div>

          <table style="width:100%;border-collapse:collapse;margin-top:10px;">
            <thead>
              <tr style="background-color:#f2f4f8;">
                <th style="padding:10px; text-align:left;">Title</th>
                <th style="padding:10px; text-align:left;">Type</th>
              </tr>
            </thead>
            <tbody>${rows}</tbody>
          </table>

          <p style="margin-top:30px; color:#555; font-size:12px; text-align:center;">
            This is an automated message from the Finance Document Management System.
          </p>
        </div>
        <div style="background-color: #eeeeee; padding: 20px; text-align: center; color: #888;">
          &copy; Finance Department - Office of the Dept Manager.
        </div>
      </div>
    </div>`;

  MailApp.sendEmail({
    to: recipientData.email,
    subject: subject,
    htmlBody: html,
    name: 'Finance Department'
  });
}
