function doGet(e) {
  // 1. Default to Dashboard
  if (!e.parameter.page && !e.parameter.action) {
    return HtmlService.createHtmlOutputFromFile('Dashboard').setTitle('Admin Dashboard - Finance');
  }

  // 2. Submit Page (External/Regular)
  if (e.parameter.page === 'submit') {
    let template = HtmlService.createTemplateFromFile('Index');
    template.docTypes = getDocTypes(); 
    template.directory = getDirectory(); 
    return template.evaluate().setTitle('Finance Document Submission');
  } 

  // 3. Internal Logging Page (Step 1 - Log & Print Slip)
  else if (e.parameter.page === 'internal-log') {
    let template = HtmlService.createTemplateFromFile('InternalLogForm');
    template.docTypes = getDocTypes();
    template.directory = getDirectory();
    template.prefillDocs = e.parameter.prefill || "";
    return template.evaluate().setTitle('Internal Document Logger');
  }

  // 4. Internal Update Page (NEW: Step 2 - Scan QR & Pass Custody)
  else if (e.parameter.page === 'internal-update') {
    const logId = e.parameter.id;
    const logDetails = getInternalLogDetails(logId); // Fetches the batch details
    
    let template = HtmlService.createTemplateFromFile('InternalUpdate');
    template.logId = logId;
    template.details = logDetails; 
    template.directory = getDirectory(); // For autocomplete
    
    return template.evaluate().setTitle('Update Custody - ' + logId);
  }
  
  // 5. Update/Manage Page (Main Workflow)
  else if (e.parameter.action === 'update') {
    const transactionId = e.parameter.id;
    const transactionDetails = getTransactionDetails(transactionId);
    
    let template = HtmlService.createTemplateFromFile('Update');
    template.transactionId = transactionId;
    template.details = transactionDetails; 
    template.directory = getDirectory();
    
    return template.evaluate().setTitle('Update Transaction - ' + transactionId);
  } 
  
  // 6. Print Slip (Main Workflow)
  else if (e.parameter.action === 'print-slip') {
    const transactionId = e.parameter.id;
    const transactionDetails = getTransactionDetails(transactionId);
    
    let template = HtmlService.createTemplateFromFile('PrintSlip');
    template.transactionId = transactionId;
    template.details = transactionDetails;
    template.appUrl = ScriptApp.getService().getUrl();
    
    return template.evaluate().setTitle('Print Slip - ' + transactionId);
  }

  // 7. Internal Print Slip (NEW: For Reprinting)
  else if (e.parameter.action === 'internal-print-slip') {
    const logId = e.parameter.id;
    const logDetails = getInternalLogDetails(logId);
    
    let template = HtmlService.createTemplateFromFile('InternalPrintSlip');
    template.logId = logId;
    template.details = logDetails;
    template.appUrl = ScriptApp.getService().getUrl();
    
    return template.evaluate().setTitle('Print Slip - ' + logId);
  }
  
  return HtmlService.createHtmlOutputFromFile('Dashboard');
}

// SAFE URL GETTER
function getWebAppUrlSafe() {
  try {
    return ScriptApp.getService().getUrl();
  } catch (e) {
    return ""; 
  }
}

function getDocTypes() {
  try {
    const docTypesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DocTypes');
    if (!docTypesSheet || docTypesSheet.getLastRow() < 2) return [];
    const types = docTypesSheet.getRange('A2:A' + docTypesSheet.getLastRow()).getValues();
    return types.map(function(row) { return row[0]; }).filter(function(type) { return type; }); 
  } catch (e) {
    console.error("Error fetching DocTypes: " + e.toString());
    return [];
  }
}

function generateTransactionID() {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000); 

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const transactionsSheet = ss.getSheetByName('Transactions');
    if (!transactionsSheet) { throw new Error("CRITICAL ERROR: 'Transactions' sheet not found."); }

    const today = new Date();
    const currentMonthYear = today.getFullYear() + '-' + ('0' + (today.getMonth() + 1)).slice(-2);
    
    const lastRow = transactionsSheet.getLastRow();
    let maxSerial = 0;
    
    if (lastRow > 1) { 
      const allTransactionIDs = transactionsSheet.getRange('A2:A' + lastRow).getValues().flat();
      for (let i = 0; i < allTransactionIDs.length; i++) {
        let id = String(allTransactionIDs[i]);
        if (id && id.startsWith(currentMonthYear)) {
          let serialPart = parseInt(id.substring(8), 10);
          if (serialPart > maxSerial) {
            maxSerial = serialPart;
          }
        }
      }
    }
    
    const newSerial = maxSerial + 1;
    const serialFormatted = ('000' + newSerial).slice(-3);
    return currentMonthYear + '-' + serialFormatted;

  } finally {
    lock.releaseLock();
  }
}

function processForm(formObject) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const transactionsSheet = ss.getSheetByName('Transactions');
  const documentsSheet = ss.getSheetByName('Documents');

  if (!transactionsSheet || !documentsSheet) {
    throw new Error("CRITICAL ERROR: Sheets not found.");
  }
  
  const newTransactionID = generateTransactionID();
  const timestamp = new Date();
  const txDetails = formObject.transactionDetails;
  
  // 1. Log Transaction Header
  transactionsSheet.appendRow([
    newTransactionID, 
    txDetails.contactPerson, 
    txDetails.centerDept, 
    txDetails.officeUnit, 
    txDetails.email, 
    timestamp
  ]);
  
  const initialStatus = "Pending Signature";
  const documentsForEmail = []; 
  const documentsWithOwner = []; 

  // 2. Log Document Items (Shifted Columns - Removed Date)
  formObject.documents.forEach(function(doc) {
    const principalName = doc.principalName || "";
    const principalEmail = doc.principalEmail || "";

    documentsSheet.appendRow([
      newTransactionID,
      doc.docType,
      doc.docTitle,
      // Removed Dates Column Here
      initialStatus, // D
      timestamp,     // E
      "",            // F (Sig Link)
      "",            // G (Comments)
      principalName, // H
      principalEmail // I
    ]);

    const docData = {
      title: doc.docTitle,
      type: doc.docType,
      owner: principalName || "(Liaison)", 
      principalEmail: principalEmail
    };

    documentsForEmail.push(docData);
    documentsWithOwner.push(docData);
  });

  try {
    sendLiaisonReceiptEmail(txDetails, newTransactionID, documentsForEmail);

    const ownerGroups = {};
    documentsWithOwner.forEach(doc => {
        if (doc.principalEmail && doc.principalEmail !== txDetails.email) {
            if (!ownerGroups[doc.principalEmail]) {
                ownerGroups[doc.principalEmail] = { name: doc.owner, docs: [] };
            }
            ownerGroups[doc.principalEmail].docs.push(doc);
        }
    });

    Object.keys(ownerGroups).forEach(email => {
        const group = ownerGroups[email];
        sendOwnerReceiptEmail(group.name, email, newTransactionID, group.docs, txDetails.contactPerson);
    });

  } catch(e) {
    console.error("Email failed: " + e.toString());
  }

  return { 
    transactionId: newTransactionID, 
    appUrl: getWebAppUrlSafe(),
    details: {
      ContactPerson: txDetails.contactPerson,
      Timestamp: timestamp.toLocaleString('en-US', { timeZone: 'Asia/Manila' }),
      documents: documentsForEmail 
    }
  };
}

// 1. LIAISON RECEIPT
function sendLiaisonReceiptEmail(txDetails, transactionId, documents) {
  const subject = `[Finance] Transaction Logged - Ref: ${transactionId}`;
  
  let documentsHtmlList = documents.map(doc => 
    `<tr>
       <td style="padding: 12px; border-bottom: 1px solid #eee;"><b>${doc.owner}</b></td>
       <td style="padding: 12px; border-bottom: 1px solid #eee;">${doc.title}</td>
       <td style="padding: 12px; border-bottom: 1px solid #eee;">${doc.type}</td>
     </tr>`
  ).join('');

  const emailBody = `
    <div style="background-color: #f4f6f8; padding: 40px 0; font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif;">
      <div style="max-width: 650px; margin: 0 auto; background-color: #ffffff; border-radius: 8px; overflow: hidden; box-shadow: 0 4px 10px rgba(0,0,0,0.05);">
        <div style="background-color: #1C2790; padding: 30px; text-align: center;">
          <img src="https://i.imgur.com/jaEbfAR.png" alt="DAP Logo" style="width: 400px; display: block; margin: 0 auto;">
        </div>
        <div style="padding: 40px; border-top: 6px solid #CDAE2C;">
          <h1 style="color: #1C2790; margin: 0 0 20px 0; text-align: center; font-size: 24px; text-transform: uppercase;">TRANSACTION LOGGED</h1>
          <div style="background-color: #f0f4ff; padding: 20px; text-align: center; margin: 0 0 30px 0; border-radius: 4px; border: 1px dashed #1C2790;">
            <span style="font-size: 11px; color: #555; text-transform: uppercase; display: block; margin-bottom: 5px;">Transaction Reference</span>
            <span style="font-size: 24px; font-weight: bold; color: #1C2790;">${transactionId}</span>
          </div>
          <p style="font-size: 16px; color: #333;">Dear <b>${txDetails.contactPerson}</b>,</p>
          <p style="font-size: 16px; color: #333;">We have received the documents you submitted to the Finance Office. Here is the summary:</p>
          <table style="width: 100%; border-collapse: collapse; font-size: 14px; margin-top: 20px;">
            <thead style="background-color: #f8f9fa;">
              <tr>
                <th style="padding: 12px; text-align: left;">Owner</th>
                <th style="padding: 12px; text-align: left;">Title</th>
                <th style="padding: 12px; text-align: left;">Type</th>
              </tr>
            </thead>
            <tbody>${documentsHtmlList}</tbody>
          </table>
        </div>
        <div style="background-color: #eeeeee; padding: 20px; text-align: center; font-size: 12px; color: #888;">
          <p>&copy; ${new Date().getFullYear()} Finance Department - Office of the Department Manager.</p>
        </div>
      </div>
    </div>`;
    
  MailApp.sendEmail({ to: txDetails.email, subject: subject, htmlBody: emailBody, name: 'Finance - Office of the Dept Manager' });
}

// 2. OWNER RECEIPT
function sendOwnerReceiptEmail(ownerName, ownerEmail, transactionId, documents, liaisonName) {
  const subject = `[Finance] Document Received - Ref: ${transactionId}`;
  
  let documentsHtmlList = documents.map(doc => 
    `<tr>
       <td style="padding: 12px; border-bottom: 1px solid #eee;">${doc.title}</td>
       <td style="padding: 12px; border-bottom: 1px solid #eee;">${doc.type}</td>
     </tr>`
  ).join('');

  const emailBody = `
    <div style="background-color: #f4f6f8; padding: 40px 0; font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif;">
      <div style="max-width: 650px; margin: 0 auto; background-color: #ffffff; border-radius: 8px; overflow: hidden; box-shadow: 0 4px 10px rgba(0,0,0,0.05);">
        <div style="background-color: #1C2790; padding: 30px; text-align: center;">
          <img src="https://i.imgur.com/jaEbfAR.png" alt="DAP Logo" style="width: 400px; display: block; margin: 0 auto;">
        </div>
        <div style="padding: 40px; border-top: 6px solid #CDAE2C;">
          <h1 style="color: #1C2790; margin: 0 0 20px 0; text-align: center; font-size: 24px; text-transform: uppercase;">DOCUMENT RECEIVED</h1>
          <div style="background-color: #f0f4ff; padding: 20px; text-align: center; margin: 0 0 30px 0; border-radius: 4px; border: 1px dashed #1C2790;">
            <span style="font-size: 11px; color: #555; text-transform: uppercase; display: block; margin-bottom: 5px;">Transaction Reference</span>
            <span style="font-size: 24px; font-weight: bold; color: #1C2790;">${transactionId}</span>
          </div>
          <p style="font-size: 16px; color: #333;">Dear <b>${ownerName}</b>,</p>
          <p style="font-size: 16px; color: #333;">This email confirms that the Finance Office has received the following document(s) submitted on your behalf by <b>${liaisonName}</b>:</p>
          <table style="width: 100%; border-collapse: collapse; font-size: 14px; margin-top: 20px;">
            <thead style="background-color: #f8f9fa;">
              <tr>
                <th style="padding: 12px; text-align: left;">Title</th>
                <th style="padding: 12px; text-align: left;">Type</th>
              </tr>
            </thead>
            <tbody>${documentsHtmlList}</tbody>
          </table>
          <p style="margin-top: 30px; font-size: 14px; color: #666; font-style: italic;">You will receive another notification once your document is processed.</p>
        </div>
        <div style="background-color: #eeeeee; padding: 20px; text-align: center; font-size: 12px; color: #888;">
          <p>&copy; ${new Date().getFullYear()} Finance Department - Office of the Department Manager.</p>
        </div>
      </div>
    </div>`;
    
  MailApp.sendEmail({ to: ownerEmail, subject: subject, htmlBody: emailBody, name: 'Finance - Office of the Dept Manager' });
}

function getDashboardData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const transactionsSheet = ss.getSheetByName('Transactions');
    const documentsSheet = ss.getSheetByName('Documents');
    
    // NEW SHEETS
    const intTxSheet = ss.getSheetByName('InternalTransactions'); 
    const intDocSheet = ss.getSheetByName('InternalDocuments');

    if (!transactionsSheet || !documentsSheet) { 
      return { transactions: [], docTypes: [], appUrl: getWebAppUrlSafe(), internalLog: [] }; 
    }

    // --- MAIN WORKFLOW DATA (Existing Logic) ---
    const docsByTxId = {};
    // ... (Your existing logic for Main Docs, kept separate for brevity but assume it is here) ...
    if (documentsSheet.getLastRow() > 1) {
      const docRange = documentsSheet.getRange(2, 1, documentsSheet.getLastRow() - 1, documentsSheet.getLastColumn());
      const docValues = docRange.getValues();
      const docHeaders = documentsSheet.getRange(1, 1, 1, documentsSheet.getLastColumn()).getValues()[0];
      docValues.forEach(function(row, index) {
        if (!row[0]) return; 
        let docObject = { rowNumber: index + 2 }; 
        docHeaders.forEach(function(header, i) { 
          if(header) {
            const cleanHeader = header.trim(); 
            if (row[i] instanceof Date) {
               docObject[cleanHeader] = row[i].toLocaleString('en-US', { timeZone: 'Asia/Manila' });
            } else {
               docObject[cleanHeader] = String(row[i]); 
            }
          }
        });
        const txId = docObject.TransactionID;
        if (txId) {
          if (!docsByTxId[txId]) { docsByTxId[txId] = []; }
          docsByTxId[txId].push(docObject);
        }
      });
    }

    const allTransactions = [];
    const docTypes = new Set();
    
    if (transactionsSheet.getLastRow() > 1) {
      const txRange = transactionsSheet.getRange(2, 1, transactionsSheet.getLastRow() - 1, transactionsSheet.getLastColumn());
      const txValues = txRange.getValues();
      const txHeaders = transactionsSheet.getRange(1, 1, 1, transactionsSheet.getLastColumn()).getValues()[0];

      txValues.forEach(function(row) {
        if (!row[0]) return;
        let txObject = {};
        txHeaders.forEach(function(header, i) {
          if(!header) return;
          const cleanHeader = header.trim(); 
          if (row[i] instanceof Date) {
            txObject[cleanHeader] = row[i].toLocaleString('en-US', { timeZone: 'Asia/Manila' });
          } else {
            txObject[cleanHeader] = String(row[i]);
          }
        });
        if (txObject.TransactionID) {
          const relatedDocs = docsByTxId[txObject.TransactionID] || [];
          txObject.documents = relatedDocs;
          relatedDocs.forEach(function(d) { if(d.DocumentType) docTypes.add(d.DocumentType); });
          allTransactions.push(txObject);
        }
      });
    }

    // --- INTERNAL LOGBOOK LOGIC (Refactored) ---
    const internalLog = [];
    
    // 1. Fetch Transactions (Timestamps)
    const txMap = {}; // Map LogID -> Timestamp
    if (intTxSheet && intTxSheet.getLastRow() > 1) {
       const txData = intTxSheet.getRange(2, 1, intTxSheet.getLastRow()-1, 2).getValues(); // Cols A & B
       txData.forEach(r => {
          if(r[0]) txMap[String(r[0])] = (r[1] instanceof Date) ? r[1].toLocaleString('en-US', { timeZone: 'Asia/Manila' }) : String(r[1]);
       });
    }

    // 2. Fetch Documents & Join
    if (intDocSheet && intDocSheet.getLastRow() > 1) {
       // Cols: LogID(0), Title(1), Type(2), Status(3), Custody(4), History(5)
       const docData = intDocSheet.getRange(2, 1, intDocSheet.getLastRow()-1, 6).getValues();
       
       docData.forEach(r => {
          if(!r[0]) return;
          // Add to docTypes for filter
          if(r[2]) docTypes.add(String(r[2]));

          internalLog.push({
             LogID: String(r[0]),
             Date: txMap[String(r[0])] || "", // Join timestamp from Tx sheet
             Title: String(r[1]),
             Type: String(r[2]),
             Status: String(r[3]), // e.g. Finance Custody, Routed
             RoutedTo: String(r[4]), // Using 'RoutedTo' key to match your Dashboard.html
             Remarks: String(r[5])   // History/Notes
          });
       });
    }
    
    return {
      transactions: allTransactions.reverse(),
      docTypes: Array.from(docTypes),
      appUrl: getWebAppUrlSafe(),
      internalLog: internalLog.reverse(),
      directory: getDirectory()
    };

  } catch (e) {
    console.error("getDashboardData CRASHED: " + e.toString());
    return { transactions: [], docTypes: [], appUrl: getWebAppUrlSafe(), internalLog: [], directory: [] };
  }
}

function getTransactionDetails(transactionId) {
  try {
    if (!transactionId) return null;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const transactionsSheet = ss.getSheetByName('Transactions');
    const documentsSheet = ss.getSheetByName('Documents');
    
    let transactionInfo = null;
    
    if (transactionsSheet.getLastRow() > 1) {
      const txData = transactionsSheet.getDataRange().getValues();
      const txHeaders = txData.shift();
      
      for (let i = 0; i < txData.length; i++) {
        if (String(txData[i][0]) === String(transactionId)) { 
          transactionInfo = {};
          txHeaders.forEach(function(header, idx) { 
            if(header) {
                const cleanHeader = header.trim();
                const cellValue = txData[i][idx];
                if (cellValue instanceof Date) {
                    transactionInfo[cleanHeader] = cellValue.toLocaleString('en-US', { timeZone: 'Asia/Manila' });
                } else {
                    transactionInfo[cleanHeader] = String(cellValue);
                }
            }
          });
          break;
        }
      }
    }
    
    if (!transactionInfo) return null;

    const relatedDocuments = [];
    if (documentsSheet.getLastRow() > 1) {
      const docData = documentsSheet.getDataRange().getValues();
      const docHeaders = docData.shift();
      
      for (let i = 0; i < docData.length; i++) {
        if (String(docData[i][0]) === String(transactionId)) { 
          let docObject = {};
          docHeaders.forEach(function(header, idx) { 
            if(header) {
                const cleanHeader = header.trim(); 
                const cellValue = docData[i][idx];
                if (cellValue instanceof Date) {
                   docObject[cleanHeader] = cellValue.toLocaleString('en-US', { timeZone: 'Asia/Manila' });
                } else {
                   docObject[cleanHeader] = String(cellValue);
                }
            }
          });
          docObject.rowNumber = i + 2; 
          relatedDocuments.push(docObject);
        }
      }
    }
    
    transactionInfo.documents = relatedDocuments;
    return transactionInfo;
  } catch (e) {
    console.error("Error in getTransactionDetails: " + e.toString());
    return null;
  }
}

function updateDocumentStatus(updateData) {
  const documentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Documents');
  const lastRow = documentsSheet.getLastRow();
  const affectedDocs = []; 
  
  if (updateData.isBatch) {
    if (lastRow < 2) return "No documents to update.";
    const data = documentsSheet.getRange(2, 1, lastRow - 1, 1).getValues();
    
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]) === String(updateData.transactionId)) {
        const currentRow = i + 2;
        const currentStatus = documentsSheet.getRange(currentRow, 4).getValue(); 
        const isActive = (currentStatus !== 'Claimed/Released' && currentStatus !== 'Pulled Out');
        const isRelogging = (updateData.newStatus === 'Pending Signature');

        if (isActive || isRelogging) {
           documentsSheet.getRange(currentRow, 4).setValue(updateData.newStatus); 
           documentsSheet.getRange(currentRow, 5).setValue(new Date()); 
           if (updateData.notes) documentsSheet.getRange(currentRow, 7).setValue(updateData.notes); 
           
           affectedDocs.push({
             Title: documentsSheet.getRange(currentRow, 3).getValue(),
             Type: documentsSheet.getRange(currentRow, 2).getValue(),
             Status: updateData.newStatus,
             PrincipalName: documentsSheet.getRange(currentRow, 8).getValue(),
             PrincipalEmail: documentsSheet.getRange(currentRow, 9).getValue()
           });
        }
      }
    }
  } 
  else {
    try {
      const row = parseInt(updateData.documentRow);
      documentsSheet.getRange(row, 4).setValue(updateData.newStatus); 
      documentsSheet.getRange(row, 5).setValue(new Date()); 
      if (updateData.notes) documentsSheet.getRange(row, 7).setValue(updateData.notes); 
      
      affectedDocs.push({
         Title: documentsSheet.getRange(row, 3).getValue(),
         Type: documentsSheet.getRange(row, 2).getValue(),
         Status: updateData.newStatus,
         PrincipalName: documentsSheet.getRange(row, 8).getValue(),
         PrincipalEmail: documentsSheet.getRange(row, 9).getValue()
      });
    } catch (e) { return "Error: " + e.message; }
  }

  if (affectedDocs.length > 0) {
     const txDetails = getTransactionDetails(updateData.transactionId);
     const principalGroups = {};
     affectedDocs.forEach(doc => {
       if (doc.PrincipalEmail && doc.PrincipalEmail.trim() !== "") {
         if (!principalGroups[doc.PrincipalEmail]) {
           principalGroups[doc.PrincipalEmail] = { Name: doc.PrincipalName, Docs: [] };
         }
         principalGroups[doc.PrincipalEmail].Docs.push(doc);
       }
     });

     // NOTIFICATIONS
     Object.keys(principalGroups).forEach(email => {
       const group = principalGroups[email];
       const pDetails = { TransactionID: txDetails.TransactionID, ContactPerson: group.Name, ContactEmail: email };

       if (updateData.newStatus === 'Signed') {
          sendReadyForPickupEmail(pDetails, group.Docs, updateData.notes, false); 
       } else if (updateData.newStatus === 'For pick up, but with comments') {
          sendPickupWithCommentsEmail(pDetails, group.Docs, updateData.notes, false);
       } else if (updateData.newStatus === 'Pending Signature') {
          // NEW: Notify Principal if relogged
          sendRelogEmail(pDetails, group.Docs, updateData.notes, false);
       }
     });

     // Notify Liaison
     if (updateData.newStatus === 'Signed') {
        sendReadyForPickupEmail(txDetails, affectedDocs, updateData.notes, true);
     } else if (updateData.newStatus === 'For pick up, but with comments') {
        sendPickupWithCommentsEmail(txDetails, affectedDocs, updateData.notes, true);
     } else if (updateData.newStatus === 'Pending Signature') {
        // NEW: Notify Liaison if relogged
        sendRelogEmail(txDetails, affectedDocs, updateData.notes, true);
     }

     return `Success: Updated ${affectedDocs.length} document(s).`;
  }
  return "No documents were updated.";
}

function sendReadyForPickupEmail(txDetails, docsList, notes, isBatch) {
  const subject = `[Finance] Ready for Pickup - Ref: ${txDetails.TransactionID}`;
  const rows = docsList.map(d => `<tr><td style="padding:12px;border-bottom:1px solid #eee;">${d.Title}</td><td style="padding:12px;border-bottom:1px solid #eee;color:#555;">${d.Type}</td></tr>`).join('');
  let remarksHtml = '';
  if (notes) {
    remarksHtml = `<div style="margin-top:25px;padding:20px;background-color:#f0fcf4;border-left:5px solid #198754;"><h4 style="margin:0;color:#198754;">Remarks</h4><p style="margin:0;">${notes}</p></div>`;
  }

  const html = `
    <div style="background-color: #f4f6f8; padding: 40px 0; font-family: Arial, sans-serif;">
      <div style="max-width: 650px; margin: 0 auto; background-color: #ffffff; border-radius: 8px; overflow: hidden;">
        <div style="background-color: #1C2790; padding: 30px; text-align: center;"><img src="https://i.imgur.com/jaEbfAR.png" width="400"></div>
        <div style="padding: 40px; border-top: 6px solid #198754;">
          <h1 style="color: #198754; text-align: center;">READY FOR PICKUP</h1>
          <p>Dear <b>${txDetails.ContactPerson}</b>,</p>
          <p>The Finance Office has signed the following documents:</p>
          <table style="width:100%;border-collapse:collapse;margin-top:20px;"><thead><tr style="background-color:#198754;color:white;"><th style="padding:10px;">Title</th><th style="padding:10px;">Type</th></tr></thead><tbody>${rows}</tbody></table>
          ${remarksHtml}
          <p style="margin-top:30px;text-align:center;">You may now proceed to the Finance Office to claim your documents or have someone claim it for you.</p>
        </div>
        <div style="background-color: #eeeeee; padding: 20px; text-align: center; color: #888;">&copy; Finance Department - Office of the Dept Manager.</div>
      </div>
    </div>`;
  
  MailApp.sendEmail({ to: txDetails.ContactEmail, subject: subject, htmlBody: html, name: 'Finance - Office of the Dept Manager' });
}

function sendPickupWithCommentsEmail(txDetails, docsList, notes, isBatch) {
  const subject = `[Finance] Action Required - Ref: ${txDetails.TransactionID}`;
  const rows = docsList.map(d => `<tr><td style="padding:12px;border-bottom:1px solid #e0e0e0;">${d.Title}</td><td style="padding:12px;border-bottom:1px solid #e0e0e0;">${d.Type}</td></tr>`).join('');
  
  const html = `
    <div style="background-color: #f4f6f8; padding: 40px 0; font-family: Arial, sans-serif;">
      <div style="max-width: 650px; margin: 0 auto; background-color: #ffffff; border-radius: 8px; overflow: hidden;">
        <div style="background-color: #1C2790; padding: 30px; text-align: center;"><img src="https://i.imgur.com/jaEbfAR.png" width="400"></div>
        <div style="padding: 40px; border-top: 6px solid #ffc107;">
          <h1 style="color: #bfa006; text-align: center;">ACTION REQUIRED</h1>
          <p>Dear <b>${txDetails.ContactPerson}</b>,</p>
          <p>The Finance Office has reviewed your documents and they require attention:</p>
          <table style="width:100%;border-collapse:collapse;margin-top:20px;"><thead><tr style="background-color:#fff3cd;color:#856404;"><th style="padding:10px;">Title</th><th style="padding:10px;">Type</th></tr></thead><tbody>${rows}</tbody></table>
          <div style="margin-top:25px;padding:20px;background-color:#fff9db;border:1px solid #ffeeba;color:#856404;"><b>Comments:</b> ${notes}</div>
        </div>
        <div style="background-color: #eeeeee; padding: 20px; text-align: center; color: #888;">&copy; Finance Department - Office of the Dept Manager.</div>
      </div>
    </div>`;
  
  MailApp.sendEmail({ to: txDetails.ContactEmail, subject: subject, htmlBody: html, name: 'Finance - Office of the Dept Manager' });
}

function processClaim(claimData) {
  try {
    const documentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Documents');
    const claimTimestamp = new Date();
    const affectedDocs = [];

    // 1. DIRECTORY LOOKUP
    const directory = getDirectory();
    const claimantObj = directory.find(d => d.name === claimData.claimedBy);
    const claimantEmail = claimantObj ? claimantObj.email : "";

    // 2. TRACKING FLAG (New)
    // Tracks if the claimant has received at least one email in this transaction
    let claimantWasNotified = false;

    const lastRow = documentsSheet.getLastRow();
    const data = documentsSheet.getRange(2, 1, lastRow - 1, 1).getValues();

    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]) === String(claimData.transactionId)) {
        const row = i + 2;
        const isTarget = claimData.isBatch || (String(row) === String(claimData.documentRow));
        const currentStatus = documentsSheet.getRange(row, 4).getValue();

        if (isTarget && (currentStatus === 'Signed' || currentStatus === 'For pick up, but with comments')) {
           documentsSheet.getRange(row, 4).setValue('Claimed/Released');
           documentsSheet.getRange(row, 5).setValue(claimTimestamp);

           const existingNotes = documentsSheet.getRange(row, 7).getValue();
           const newNote = existingNotes ? `${existingNotes} [Claimed by: ${claimData.claimedBy}]` : `[Claimed by: ${claimData.claimedBy}]`;
           documentsSheet.getRange(row, 7).setValue(newNote);
           
           const pName = documentsSheet.getRange(row, 8).getValue();
           const pEmail = documentsSheet.getRange(row, 9).getValue();

           affectedDocs.push({
             Title: documentsSheet.getRange(row, 3).getValue(),
             Type: documentsSheet.getRange(row, 2).getValue(),
             Status: 'Claimed/Released',
             PrincipalName: pName,
             PrincipalEmail: pEmail
           });
        }
      }
    }

    if (affectedDocs.length > 0) {
      const txDetails = getTransactionDetails(claimData.transactionId);
      const principalGroups = {};
      
      affectedDocs.forEach(doc => {
        if (doc.PrincipalEmail && doc.PrincipalEmail.trim() !== "") {
          if (!principalGroups[doc.PrincipalEmail]) {
            principalGroups[doc.PrincipalEmail] = { Name: doc.PrincipalName, Docs: [] };
          }
          principalGroups[doc.PrincipalEmail].Docs.push(doc);
        }
      });

      // A. Notify Owners (with CC to Claimant)
      Object.keys(principalGroups).forEach(email => {
        // Prevent duplicate if Owner IS the Claimant
        if (email !== claimantEmail) {
            const group = principalGroups[email];
            const principalTxDetails = {
              TransactionID: txDetails.TransactionID,
              ContactPerson: group.Name || "Document Owner",
              ContactEmail: email
            };
            
            // Send email and Mark claimant as notified
            sendClaimedEmail(principalTxDetails, group.Docs, claimTimestamp, claimData.claimedBy, claimantEmail);
            claimantWasNotified = true;
        }
      });

      // B. Notify Liaison (Conditional CC to Claimant)
      // LOGIC FIX: If claimant hasn't been notified yet (e.g. they are a third party 
      // or the Owner loop was skipped), CC them here.
      // Also ensure we don't CC them if they ARE the Liaison (avoid self-CC).
      
      let liaisonCc = null;
      if (!claimantWasNotified && claimantEmail && claimantEmail !== txDetails.ContactEmail) {
          liaisonCc = claimantEmail;
      }

      sendClaimedEmail(txDetails, affectedDocs, claimTimestamp, claimData.claimedBy, liaisonCc);

      return `Success: Released ${affectedDocs.length} document(s).`;
    } 
    return "No eligible documents found to release.";

  } catch (e) { return "Error: " + e.message; }
}

/**
 * Sends a "Claimed/Released" notification email.
 * Theme: Standard Finance Blue. Signature-free: confirmation is name-based.
 * Logic: CC behavior is now controlled by the caller (processClaim).
 */
function sendClaimedEmail(txDetails, docsList, timestamp, claimantName, ccEmail) {
  const subject = `[Finance] Document Claimed - Ref: ${txDetails.TransactionID}`;
  const formattedTimestamp = timestamp.toLocaleString('en-US', { timeZone: 'Asia/Manila' });

  const rows = docsList.map(d =>
    `<tr>
       <td style="padding:12px;border-bottom:1px solid #eee;">${d.Title}</td>
       <td style="padding:12px;border-bottom:1px solid #eee;color:#555;">${d.Type}</td>
     </tr>`
  ).join('');

  // Use the passed ccEmail if available; otherwise blank
  const finalCc = ccEmail || "";

  const html = `
    <div style="background-color: #f4f6f8; padding: 40px 0; font-family: Arial, sans-serif;">
      <div style="max-width: 650px; margin: 0 auto; background-color: #ffffff; border-radius: 8px; overflow: hidden;">
        <div style="background-color: #1C2790; padding: 30px; text-align: center;">
          <img src="https://i.imgur.com/jaEbfAR.png" width="400">
        </div>
        <div style="padding: 40px; border-top: 6px solid #1C2790;">
          <h1 style="color: #1C2790; text-align: center;">TRANSACTION COMPLETE</h1>
          <p>Dear <b>${txDetails.ContactPerson}</b>,</p>
          <p>This email confirms that the following Finance documents were successfully claimed:</p>

          <table style="width:100%;border-collapse:collapse;margin-top:20px;">
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
            <p style="margin:8px 0; font-size:18px; font-weight:bold; color:#1C2790;">${claimantName}</p>
            <p style="font-size:12px; color:#777;">Claimed on: ${formattedTimestamp}</p>
          </div>

          <div style="margin-top:20px;padding:15px 20px;background-color:#fff9db;border:1px solid #ffeeba;color:#856404;font-size:13px;border-radius:4px;">
            <b>No reply needed.</b> If you did <u>not</u> pick up these documents, or did not authorize <b>${claimantName}</b> to pick them up on your behalf, please reply to this email to dispute it. Otherwise, no action is required &mdash; this serves as your confirmation of receipt.
          </div>
        </div>
        <div style="background-color: #eeeeee; padding: 20px; text-align: center; color: #888;">
          &copy; Finance Department - Office of the Dept Manager.
        </div>
      </div>
    </div>`;

  MailApp.sendEmail({
    to: txDetails.ContactEmail,
    cc: finalCc,
    subject: subject,
    htmlBody: html,
    name: 'Finance - Office of the Dept Manager'
  });
}

function processPullOut(pullOutData) {
  try {
    const documentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Documents');
    const timestamp = new Date();
    const affectedDocs = [];

    // 1. DIRECTORY LOOKUP
    const directory = getDirectory();
    const pullerName = pullOutData.pulledOutBy || "Unknown"; 
    const pullerObj = directory.find(d => d.name === pullerName);
    const pullerEmail = pullerObj ? pullerObj.email : "";

    const lastRow = documentsSheet.getLastRow();
    const data = documentsSheet.getRange(2, 1, lastRow - 1, 1).getValues();

    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]) === String(pullOutData.transactionId)) {
        const row = i + 2;
        const isTarget = pullOutData.isBatch || (String(row) === String(pullOutData.documentRow));
        const currentStatus = documentsSheet.getRange(row, 4).getValue();

        if (isTarget && currentStatus !== 'Claimed/Released' && currentStatus !== 'Pulled Out') {
           documentsSheet.getRange(row, 4).setValue('Pulled Out');
           documentsSheet.getRange(row, 5).setValue(timestamp);

           const fullComment = `${pullOutData.comments} [Pulled out by: ${pullerName}]`;
           documentsSheet.getRange(row, 7).setValue(fullComment);
           
           const pName = documentsSheet.getRange(row, 8).getValue();
           const pEmail = documentsSheet.getRange(row, 9).getValue();

           affectedDocs.push({
             Title: documentsSheet.getRange(row, 3).getValue(),
             Type: documentsSheet.getRange(row, 2).getValue(),
             Status: 'Pulled Out',
             PrincipalName: pName,
             PrincipalEmail: pEmail
           });
        }
      }
    }

    if (affectedDocs.length > 0) {
      const txDetails = getTransactionDetails(pullOutData.transactionId);
      const principalGroups = {};
      affectedDocs.forEach(doc => {
        if (doc.PrincipalEmail && doc.PrincipalEmail.trim() !== "") {
          if (!principalGroups[doc.PrincipalEmail]) {
            principalGroups[doc.PrincipalEmail] = { Name: doc.PrincipalName, Docs: [] };
          }
          principalGroups[doc.PrincipalEmail].Docs.push(doc);
        }
      });

      // A. Notify Owners (CC the Puller)
      Object.keys(principalGroups).forEach(email => {
        // Prevent sending to owner if owner IS the puller (they know they pulled it)
        if (email !== pullerEmail) { 
            const group = principalGroups[email];
            const principalTxDetails = {
              TransactionID: txDetails.TransactionID,
              ContactPerson: group.Name || "Document Owner",
              ContactEmail: email
            };
            
            // PASSING pullerEmail AS CC HERE
            sendPulledOutEmail(principalTxDetails, group.Docs, timestamp, pullOutData.comments, pullerName, pullerEmail);
        }
      });

      // B. Notify Liaison (NO CC)
      // We do not CC the puller here to avoid double emailing (Bombardment prevention).
      // Liaison needs the record.
      if (txDetails.ContactEmail !== pullerEmail) {
         // Pass null for CC
         sendPulledOutEmail(txDetails, affectedDocs, timestamp, pullOutData.comments, pullerName, null);
      }

      return `Success: Pulled out ${affectedDocs.length} document(s).`;
    }
    return "No eligible documents found to pull out.";

  } catch (e) { return "Error: " + e.message; }
}

/**
 * Sends a "Pulled Out" notification email.
 * Theme: Red/Warning colors. Signature-free: confirmation is name-based.
 * Update: Now accepts 'ccEmail' to copy the puller.
 */
function sendPulledOutEmail(txDetails, docsList, timestamp, reason, pulledOutByName, ccEmail) {
  const subject = `[Finance] Document Retrieved - Ref: ${txDetails.TransactionID}`;
  const formattedTimestamp = timestamp.toLocaleString('en-US', { timeZone: 'Asia/Manila' });

  const rows = docsList.map(d =>
    `<tr>
       <td style="padding:12px;border-bottom:1px solid #f8d7da;">${d.Title}</td>
       <td style="padding:12px;border-bottom:1px solid #f8d7da;color:#721c24;">${d.Type}</td>
     </tr>`
  ).join('');

  // Handle undefined/null CC
  const finalCc = ccEmail || "";

  const html = `
    <div style="background-color: #f4f6f8; padding: 40px 0; font-family: Arial, sans-serif;">
      <div style="max-width: 650px; margin: 0 auto; background-color: #ffffff; border-radius: 8px; overflow: hidden;">
        <div style="background-color: #1C2790; padding: 30px; text-align: center;">
          <img src="https://i.imgur.com/jaEbfAR.png" width="400">
        </div>
        <div style="padding: 40px; border-top: 6px solid #dc3545;">
          <h1 style="color: #dc3545; text-align: center;">RETRIEVAL NOTICE</h1>
          <p>Dear <b>${txDetails.ContactPerson}</b>,</p>
          <p>The following document(s) have been retrieved/pulled out from the Finance Office:</p>

          <table style="width:100%;border-collapse:collapse;margin-top:20px;">
            <thead>
              <tr style="background-color:#f8d7da;color:#721c24;">
                <th style="padding:10px; text-align:left;">Title</th>
                <th style="padding:10px; text-align:left;">Type</th>
              </tr>
            </thead>
            <tbody>${rows}</tbody>
          </table>

          <div style="margin-top:25px;padding:20px;background-color:#fff5f5;border:1px solid #f5c6cb;color:#721c24;">
            <b>Reason for Retrieval:</b><br>"${reason}"
          </div>

          <div style="margin-top:30px;border:1px solid #eee;padding:20px;text-align:center;">
             <p style="font-weight:bold; color:#555;">RETRIEVED BY</p>
             <p style="margin:8px 0; font-size:18px; font-weight:bold; color:#dc3545;">${pulledOutByName || "Authorized Personnel"}</p>
             <p style="font-size:12px; color:#777;">Timestamp: ${formattedTimestamp}</p>
          </div>

          <div style="margin-top:20px;padding:15px 20px;background-color:#fff9db;border:1px solid #ffeeba;color:#856404;font-size:13px;border-radius:4px;">
            <b>No reply needed.</b> If you did <u>not</u> authorize this retrieval, please reply to this email to dispute it. Otherwise, no action is required.
          </div>
        </div>
        <div style="background-color: #eeeeee; padding: 20px; text-align: center; color: #888;">
          &copy; Finance Department - Office of the Dept Manager.
        </div>
      </div>
    </div>`;

  MailApp.sendEmail({
    to: txDetails.ContactEmail,
    cc: finalCc,
    subject: subject,
    htmlBody: html,
    name: 'Finance - Office of the Dept Manager'
  });
}

function getDirectory() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Directory');
    if (!sheet || sheet.getLastRow() < 2) return [];
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
    return data.map(row => ({ name: row[0], email: row[1] })).filter(d => d.name);
  } catch (e) {
    console.error("Error fetching directory: " + e.toString());
    return [];
  }
}

function getLogDetails(logID) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("InternalLogbook");
  if (!sheet) return null;

  const data = sheet.getDataRange().getValues();
  // Loop backwards to find the *latest* entry for this Log ID
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === String(logID)) {
      return {
        logID: data[i][0],
        title: data[i][2],
        type: data[i][3]
      };
    }
  }
  return null;
}

/**
 * Sends a notification when a document is returned to Finance Custody (Re-logged).
 * Theme: Orange/Amber (Processing/Warning).
 */
function sendRelogEmail(recipientData, docsList, notes, isLiaison) {
  const subject = `[Finance] Document Returned to Custody - Ref: ${recipientData.TransactionID}`;
  
  const rows = docsList.map(d => 
    `<tr>
       <td style="padding:12px;border-bottom:1px solid #ffeeba;">${d.Title}</td>
       <td style="padding:12px;border-bottom:1px solid #ffeeba;color:#856404;">${d.Type}</td>
     </tr>`
  ).join('');

  const html = `
    <div style="background-color: #f4f6f8; padding: 40px 0; font-family: Arial, sans-serif;">
      <div style="max-width: 650px; margin: 0 auto; background-color: #ffffff; border-radius: 8px; overflow: hidden;">
        <div style="background-color: #1C2790; padding: 30px; text-align: center;">
          <img src="https://i.imgur.com/jaEbfAR.png" width="400">
        </div>
        <div style="padding: 40px; border-top: 6px solid #fd7e14;">
          <h1 style="color: #fd7e14; text-align: center;">RETURNED TO CUSTODY</h1>
          <p>Dear <b>${recipientData.ContactPerson}</b>,</p>
          <p>This is to notify you that the following document(s) have been <b>returned to Finance custody</b> (Re-logged) and are currently being processed:</p>
          
          <table style="width:100%;border-collapse:collapse;margin-top:20px;">
            <thead>
              <tr style="background-color:#fff3cd;color:#856404;">
                <th style="padding:10px; text-align:left;">Title</th>
                <th style="padding:10px; text-align:left;">Type</th>
              </tr>
            </thead>
            <tbody>${rows}</tbody>
          </table>

          <div style="background-color:#f8f9fa; border-left:4px solid #fd7e14; padding:15px; margin-top:20px;">
            <strong>Status:</strong> Pending Signature / Processing<br>
            ${notes ? `<strong>Note:</strong> ${notes}` : ''}
          </div>
          
          <p style="margin-top:20px; color:#555;">You will be notified again via email once these documents are signed and ready for release.</p>
        </div>
        <div style="background-color: #eeeeee; padding: 20px; text-align: center; color: #888;">
          &copy; Finance Department - Office of the Dept Manager.
        </div>
      </div>
    </div>`;
  
  MailApp.sendEmail({
    to: recipientData.ContactEmail,
    subject: subject,
    htmlBody: html,
    name: 'Finance Department'
  });
}


/* =====================================================================
 * MULTI-TRANSACTION ACTIONS (Dashboard checkbox selection)
 * Each function operates on an explicit list of Documents-sheet row
 * numbers, so a single action can span several transactions at once.
 * ===================================================================== */

/**
 * Updates the status of the supplied document rows (across any number of
 * transactions). Notifications mirror the single-transaction behaviour:
 * each document owner is notified about their own documents, and the
 * liaison of each affected transaction is notified about all of theirs.
 * data: { rows: [Number], newStatus: String, notes: String }
 */
function processMultiStatusUpdate(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Documents');
    const timestamp = new Date();
    const affectedByTx = {};

    (data.rows || []).forEach(function(r) {
      const row = parseInt(r, 10);
      if (!row || row < 2) return;
      const currentStatus = sheet.getRange(row, 4).getValue();
      const isActive = (currentStatus !== 'Claimed/Released' && currentStatus !== 'Pulled Out');
      const isRelogging = (data.newStatus === 'Pending Signature');
      if (!isActive && !isRelogging) return;

      sheet.getRange(row, 4).setValue(data.newStatus);
      sheet.getRange(row, 5).setValue(timestamp);
      if (data.notes) sheet.getRange(row, 7).setValue(data.notes);

      const txId = String(sheet.getRange(row, 1).getValue());
      if (!affectedByTx[txId]) affectedByTx[txId] = [];
      affectedByTx[txId].push({
        Title: sheet.getRange(row, 3).getValue(),
        Type: sheet.getRange(row, 2).getValue(),
        Status: data.newStatus,
        PrincipalName: sheet.getRange(row, 8).getValue(),
        PrincipalEmail: sheet.getRange(row, 9).getValue()
      });
    });

    let total = 0;
    Object.keys(affectedByTx).forEach(function(txId) {
      const docs = affectedByTx[txId];
      total += docs.length;
      const txDetails = getTransactionDetails(txId);
      if (!txDetails) return;

      const principalGroups = {};
      docs.forEach(function(doc) {
        const email = String(doc.PrincipalEmail || "").trim();
        if (!email) return;
        if (!principalGroups[email]) principalGroups[email] = { Name: doc.PrincipalName, Docs: [] };
        principalGroups[email].Docs.push(doc);
      });

      Object.keys(principalGroups).forEach(function(email) {
        const g = principalGroups[email];
        const pDetails = { TransactionID: txId, ContactPerson: g.Name || "Document Owner", ContactEmail: email };
        if (data.newStatus === 'Signed') sendReadyForPickupEmail(pDetails, g.Docs, data.notes, false);
        else if (data.newStatus === 'For pick up, but with comments') sendPickupWithCommentsEmail(pDetails, g.Docs, data.notes, false);
        else if (data.newStatus === 'Pending Signature') sendRelogEmail(pDetails, g.Docs, data.notes, false);
      });

      if (data.newStatus === 'Signed') sendReadyForPickupEmail(txDetails, docs, data.notes, true);
      else if (data.newStatus === 'For pick up, but with comments') sendPickupWithCommentsEmail(txDetails, docs, data.notes, true);
      else if (data.newStatus === 'Pending Signature') sendRelogEmail(txDetails, docs, data.notes, true);
    });

    if (total === 0) return "No documents were updated. They may already be closed.";
    return `Success: Updated ${total} document(s) to "${data.newStatus}".`;
  } catch (e) {
    return "Error: " + e.message;
  } finally {
    lock.releaseLock();
  }
}

/**
 * Releases (marks as picked up) the supplied document rows to a single
 * claimant, across any number of transactions. Signature-free.
 * Email policy: each document owner gets a separate email containing only
 * their own documents; the claimant gets one consolidated email.
 * data: { rows: [Number], claimedBy: String }
 */
function processMultiClaim(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Documents');
    const timestamp = new Date();
    const directory = getDirectory();
    const claimantKey = String(data.claimedBy || "").trim().toLowerCase();
    const claimantObj = directory.find(function(d) { return String(d.name).trim().toLowerCase() === claimantKey; });
    const claimantEmail = claimantObj ? claimantObj.email : "";
    const affected = [];

    (data.rows || []).forEach(function(r) {
      const row = parseInt(r, 10);
      if (!row || row < 2) return;
      const status = sheet.getRange(row, 4).getValue();
      if (status !== 'Signed' && status !== 'For pick up, but with comments') return;

      sheet.getRange(row, 4).setValue('Claimed/Released');
      sheet.getRange(row, 5).setValue(timestamp);
      const existingNotes = sheet.getRange(row, 7).getValue();
      const note = `[Claimed by: ${data.claimedBy}]`;
      sheet.getRange(row, 7).setValue(existingNotes ? `${existingNotes} ${note}` : note);

      affected.push({
        Ref: String(sheet.getRange(row, 1).getValue()),
        Title: sheet.getRange(row, 3).getValue(),
        Type: sheet.getRange(row, 2).getValue(),
        PrincipalName: sheet.getRange(row, 8).getValue(),
        PrincipalEmail: sheet.getRange(row, 9).getValue()
      });
    });

    if (affected.length === 0) {
      return "No eligible documents found. Only documents marked 'Signed' or 'For pick up' can be picked up.";
    }

    // Each owner: a separate email with only their own documents.
    const ownerGroups = {};
    affected.forEach(function(doc) {
      const email = String(doc.PrincipalEmail || "").trim();
      if (!email || email.toLowerCase() === String(claimantEmail).toLowerCase()) return;
      if (!ownerGroups[email]) ownerGroups[email] = { name: doc.PrincipalName || "Document Owner", docs: [] };
      ownerGroups[email].docs.push(doc);
    });
    Object.keys(ownerGroups).forEach(function(email) {
      sendPickupConfirmationEmail({ name: ownerGroups[email].name, email: email }, ownerGroups[email].docs, data.claimedBy, timestamp, false);
    });

    // The claimant: one consolidated email covering everything.
    if (claimantEmail) {
      sendPickupConfirmationEmail({ name: data.claimedBy, email: claimantEmail }, affected, data.claimedBy, timestamp, true);
    }

    const txCount = Object.keys(affected.reduce(function(m, d) { m[d.Ref] = 1; return m; }, {})).length;
    return `Success: Released ${affected.length} document(s) across ${txCount} transaction(s) to ${data.claimedBy}.`;
  } catch (e) {
    return "Error: " + e.message;
  } finally {
    lock.releaseLock();
  }
}

/**
 * Pulls out the supplied document rows across any number of transactions.
 * Signature-free. Email policy mirrors processMultiClaim.
 * data: { rows: [Number], pulledOutBy: String, comments: String }
 */
function processMultiPullOut(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Documents');
    const timestamp = new Date();
    const directory = getDirectory();
    const pullerName = data.pulledOutBy || "Unknown";
    const pullerKey = String(pullerName).trim().toLowerCase();
    const pullerObj = directory.find(function(d) { return String(d.name).trim().toLowerCase() === pullerKey; });
    const pullerEmail = pullerObj ? pullerObj.email : "";
    const affected = [];

    (data.rows || []).forEach(function(r) {
      const row = parseInt(r, 10);
      if (!row || row < 2) return;
      const status = sheet.getRange(row, 4).getValue();
      if (status === 'Claimed/Released' || status === 'Pulled Out') return;

      sheet.getRange(row, 4).setValue('Pulled Out');
      sheet.getRange(row, 5).setValue(timestamp);
      const existingNotes = sheet.getRange(row, 7).getValue();
      const note = `${data.comments} [Pulled out by: ${pullerName}]`;
      sheet.getRange(row, 7).setValue(existingNotes ? `${existingNotes} ${note}` : note);

      affected.push({
        Ref: String(sheet.getRange(row, 1).getValue()),
        Title: sheet.getRange(row, 3).getValue(),
        Type: sheet.getRange(row, 2).getValue(),
        PrincipalName: sheet.getRange(row, 8).getValue(),
        PrincipalEmail: sheet.getRange(row, 9).getValue()
      });
    });

    if (affected.length === 0) return "No eligible documents found to pull out.";

    const ownerGroups = {};
    affected.forEach(function(doc) {
      const email = String(doc.PrincipalEmail || "").trim();
      if (!email || email.toLowerCase() === String(pullerEmail).toLowerCase()) return;
      if (!ownerGroups[email]) ownerGroups[email] = { name: doc.PrincipalName || "Document Owner", docs: [] };
      ownerGroups[email].docs.push(doc);
    });
    Object.keys(ownerGroups).forEach(function(email) {
      sendPullOutConfirmationEmail({ name: ownerGroups[email].name, email: email }, ownerGroups[email].docs, pullerName, timestamp, data.comments, false);
    });

    if (pullerEmail) {
      sendPullOutConfirmationEmail({ name: pullerName, email: pullerEmail }, affected, pullerName, timestamp, data.comments, true);
    }

    const txCount = Object.keys(affected.reduce(function(m, d) { m[d.Ref] = 1; return m; }, {})).length;
    return `Success: Pulled out ${affected.length} document(s) across ${txCount} transaction(s).`;
  } catch (e) {
    return "Error: " + e.message;
  } finally {
    lock.releaseLock();
  }
}

/**
 * Pickup confirmation email (signature-free). Shows the main-workflow
 * reference number per document. isPicker=true => consolidated copy for
 * the person who picked up; false => owner copy (their documents only).
 */
function sendPickupConfirmationEmail(recipient, docsList, claimantName, timestamp, isPicker) {
  try {
    if (!recipient || !recipient.email) return;
    const formattedTimestamp = timestamp.toLocaleString('en-US', { timeZone: 'Asia/Manila' });
    const subject = isPicker
      ? `[Finance] Documents Picked Up - ${docsList.length} document(s)`
      : `[Finance] Your Document Has Been Picked Up`;

    const rows = docsList.map(function(d) {
      return `<tr>
         <td style="padding:10px;border-bottom:1px solid #eee;font-weight:bold;color:#1C2790;">${d.Ref || ""}</td>
         <td style="padding:10px;border-bottom:1px solid #eee;">${d.Title}</td>
         <td style="padding:10px;border-bottom:1px solid #eee;color:#555;">${d.Type}</td>
       </tr>`;
    }).join('');

    const introText = isPicker
      ? `This email confirms that you picked up the following document(s) from the Finance Office:`
      : `This email confirms that <b>${claimantName}</b> picked up the following document(s) belonging to you from the Finance Office:`;

    const html = `
      <div style="background-color: #f4f6f8; padding: 40px 0; font-family: Arial, sans-serif;">
        <div style="max-width: 650px; margin: 0 auto; background-color: #ffffff; border-radius: 8px; overflow: hidden;">
          <div style="background-color: #1C2790; padding: 30px; text-align: center;">
            <img src="https://i.imgur.com/jaEbfAR.png" width="400">
          </div>
          <div style="padding: 40px; border-top: 6px solid #198754;">
            <h1 style="color: #198754; text-align: center;">DOCUMENTS PICKED UP</h1>
            <p>Dear <b>${recipient.name}</b>,</p>
            <p>${introText}</p>

            <table style="width:100%;border-collapse:collapse;margin-top:20px;">
              <thead>
                <tr style="background-color:#f2f4f8;">
                  <th style="padding:10px; text-align:left;">Reference #</th>
                  <th style="padding:10px; text-align:left;">Title</th>
                  <th style="padding:10px; text-align:left;">Type</th>
                </tr>
              </thead>
              <tbody>${rows}</tbody>
            </table>

            <div style="margin-top:30px;border:1px solid #eee;padding:20px;text-align:center;">
              <p style="font-weight:bold; color:#555;">RECEIVED BY</p>
              <p style="margin:8px 0; font-size:18px; font-weight:bold; color:#198754;">${claimantName}</p>
              <p style="font-size:12px; color:#777;">Picked up on: ${formattedTimestamp}</p>
            </div>

            <div style="margin-top:20px;padding:15px 20px;background-color:#fff9db;border:1px solid #ffeeba;color:#856404;font-size:13px;border-radius:4px;">
              <b>No reply needed.</b> If you did <u>not</u> pick up these documents, or did not authorize <b>${claimantName}</b> to pick them up on your behalf, please reply to this email to dispute it. Otherwise, no action is required &mdash; this serves as your confirmation of receipt.
            </div>
          </div>
          <div style="background-color: #eeeeee; padding: 20px; text-align: center; color: #888;">
            &copy; Finance Department - Office of the Dept Manager.
          </div>
        </div>
      </div>`;

    MailApp.sendEmail({
      to: recipient.email,
      subject: subject,
      htmlBody: html,
      name: 'Finance - Office of the Dept Manager'
    });
  } catch (e) {
    console.error("sendPickupConfirmationEmail failed: " + e.toString());
  }
}

/**
 * Pull-out confirmation email (signature-free). isPuller=true => the
 * person who pulled the documents out; false => owner copy.
 */
function sendPullOutConfirmationEmail(recipient, docsList, pullerName, timestamp, reason, isPuller) {
  try {
    if (!recipient || !recipient.email) return;
    const formattedTimestamp = timestamp.toLocaleString('en-US', { timeZone: 'Asia/Manila' });
    const subject = isPuller
      ? `[Finance] Documents Pulled Out - ${docsList.length} document(s)`
      : `[Finance] Your Document Has Been Pulled Out`;

    const rows = docsList.map(function(d) {
      return `<tr>
         <td style="padding:10px;border-bottom:1px solid #f8d7da;font-weight:bold;color:#721c24;">${d.Ref || ""}</td>
         <td style="padding:10px;border-bottom:1px solid #f8d7da;">${d.Title}</td>
         <td style="padding:10px;border-bottom:1px solid #f8d7da;color:#555;">${d.Type}</td>
       </tr>`;
    }).join('');

    const introText = isPuller
      ? `This email confirms that you pulled out the following document(s) from the Finance Office:`
      : `This email confirms that <b>${pullerName}</b> pulled out the following document(s) belonging to you from the Finance Office:`;

    const html = `
      <div style="background-color: #f4f6f8; padding: 40px 0; font-family: Arial, sans-serif;">
        <div style="max-width: 650px; margin: 0 auto; background-color: #ffffff; border-radius: 8px; overflow: hidden;">
          <div style="background-color: #1C2790; padding: 30px; text-align: center;">
            <img src="https://i.imgur.com/jaEbfAR.png" width="400">
          </div>
          <div style="padding: 40px; border-top: 6px solid #dc3545;">
            <h1 style="color: #dc3545; text-align: center;">DOCUMENTS PULLED OUT</h1>
            <p>Dear <b>${recipient.name}</b>,</p>
            <p>${introText}</p>

            <table style="width:100%;border-collapse:collapse;margin-top:20px;">
              <thead>
                <tr style="background-color:#f8d7da;color:#721c24;">
                  <th style="padding:10px; text-align:left;">Reference #</th>
                  <th style="padding:10px; text-align:left;">Title</th>
                  <th style="padding:10px; text-align:left;">Type</th>
                </tr>
              </thead>
              <tbody>${rows}</tbody>
            </table>

            <div style="margin-top:25px;padding:20px;background-color:#fff5f5;border:1px solid #f5c6cb;color:#721c24;">
              <b>Reason for Pull Out:</b><br>"${reason}"
            </div>

            <div style="margin-top:25px;border:1px solid #eee;padding:20px;text-align:center;">
              <p style="font-weight:bold; color:#555;">RETRIEVED BY</p>
              <p style="margin:8px 0; font-size:18px; font-weight:bold; color:#dc3545;">${pullerName}</p>
              <p style="font-size:12px; color:#777;">Pulled out on: ${formattedTimestamp}</p>
            </div>

            <div style="margin-top:20px;padding:15px 20px;background-color:#fff9db;border:1px solid #ffeeba;color:#856404;font-size:13px;border-radius:4px;">
              <b>No reply needed.</b> If you did <u>not</u> authorize this retrieval, please reply to this email to dispute it. Otherwise, no action is required.
            </div>
          </div>
          <div style="background-color: #eeeeee; padding: 20px; text-align: center; color: #888;">
            &copy; Finance Department - Office of the Dept Manager.
          </div>
        </div>
      </div>`;

    MailApp.sendEmail({
      to: recipient.email,
      subject: subject,
      htmlBody: html,
      name: 'Finance - Office of the Dept Manager'
    });
  } catch (e) {
    console.error("sendPullOutConfirmationEmail failed: " + e.toString());
  }
}
