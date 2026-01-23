# Digital Document Logbook & Management System (GAS)

**Development Academy of the Philippines (DAP)**
*Tailored for: Finance Department - Office of the Department Manager*

## 📖 Overview

This **Google Apps Script (GAS) Web Application** is designed to modernize document tracking within the Finance Department. It replaces manual logbooks with a secure, cloud-based system that ensures accountability through digital initials and real-time transparency via automated email notifications.

### Why This System Exists (The Problem with Manual Logbooks)
This system serves as a powerful alternative to **manual handwritten logbooks**, effectively eliminating common administrative woes such as:
* **Illegible Handwriting:** No more guessing who signed for a document because of scribbled names.
* **"Lost in Transit" Files:** Eliminates the "he-said-she-said" disputes when a document goes missing on someone's desk without a formal handover.
* **The Search Nightmare:** Stops the wasted hours spent flipping through physical logbook pages just to find the status of a single transaction from months ago.
* **Lack of Notification:** Removes the need for clients to physically follow up or call the office to check if their document is signed—the system emails them automatically.

By digitizing these flows with **QR Codes, Digital Initials, and Automated Emails**, this system ensures that every document is accounted for, instantly searchable, and transparently tracked from entry to release.

---

## 🚀 Workflows & Features

### 1. Document Receiving & Intake (Finance Staff Role)

* **Digital Logging Portal:** The Finance Staff (Receiver) uses the submission form (`Index.html`) to log incoming documents on behalf of the Liaison or Client.
* **Batch Processing:** Staff can log multiple documents under a single Transaction ID to speed up intake.
* **Smart QR Slips:** The system automatically generates a PDF-ready slip with a QR code (`PrintSlip.html`). The Staff prints and attaches this slip to the physical file—effectively "digitizing" it. Scanning this QR code later instantly retrieves the live digital record.
* **Automated Receipt:** Upon logging, the system automatically emails a "Receipt of Acknowledgment" to the Liaison and Document Owners, confirming that Finance has officially received their documents.

### 2. Internal Logbook Workflow (Custody Management)

* **Internal Tracking:** Tracks document movement *within* the department (e.g., routing to staff for review) using `InternalLogForm.html`.
* **"Received From" Logging:** Logs the origin of documents and sends "Proof of Receipt" emails to the source.
* **Batch Routing:** Allows passing custody of multiple documents to a colleague with a single digital initial using `InternalUpdate.html`.
* **Return to Custody:** Simplified workflow for returning documents to the central releasing officer.

### 3. Core System Capabilities

* **Admin Dashboard:** A central command center (`Dashboard.html`) to view Active, Archived, and Internal logs.
* **Digital Initials:** Integrated `signature_pad` for capturing initials on tablets/desktops.
* **Directory Integration:** Fetches names and emails from a Google Sheet to prevent typos and ensure notifications reach the right person.
* **Status Tracking:** Visual status indicators (e.g., "Pending Signature", "Signed", "Released").

---

## 🛠️ Installation & Setup

### Prerequisites

* A Google Workspace Account (DAP domain recommended).
* Access to Google Drive and Google Sheets.

### Step 1: Database Setup (Google Sheets)

Create a new Google Sheet. This will serve as your database. You must create the following tabs with the **exact** header names in Row 1:

1. **`Transactions`** (External Headers)
   * `TransactionID`, `ContactPerson`, `CenterDept`, `OfficeUnit`, `Email`, `Timestamp`

2. **`Documents`** (External Details)
   * `TransactionID`, `DocumentType`, `DocumentTitle`, `Status`, `LastUpdated`, `SignatureLink`, `Comments`, `PrincipalName`, `PrincipalEmail`

3. **`InternalTransactions`** (Internal Headers)
   * `LogID`, `Timestamp`, `Encoder`, `TotalDocs`

4. **`InternalDocuments`** (Internal Details)
   * `LogID`, `DocumentTitle`, `DocumentType`, `Status`, `CurrentCustody`, `History`, `SignatureID`

5. **`Directory`** (For Autocomplete & Emails)
   * `Name`, `Email Address`

6. **`DocTypes`** (For Dropdowns)
   * `Document Type` (List them in Column A: e.g., DV, ORS, PR, Letter)

### Step 2: Code Deployment

1. Open the Google Sheet.
2. Go to **Extensions** > **Apps Script**.
3. Copy and paste the files from this repository into the script editor:
   * `Code.gs` (Server-side logic)
   * `InternalLogging.js` (Internal logic)
   * `Dashboard.html`, `Index.html`, `Update.html`, `InternalLogForm.html`, `InternalUpdate.html`, `Track.html`, `Receipt.html`, `PrintSlip.html`, `InternalPrintSlip.html`, `Pickup.html`, `pull-out.html` (Client-side views).
4. **Important:** Ensure all filenames match exactly as they appear in the repo.

### Step 3: Initial Configuration

1. **Folder Setup:** Run the system once. It will automatically create a folder named `DMS Signatures` in your Google Drive to store initial images.
2. **Deploy:**
   * Click **Deploy** > **New Deployment**.
   * Select type: **Web app**.
   * Execute as: **Me** (The admin account).
   * Who has access: **Anyone within [Organization]** (or "Anyone" if liaisons are external).
   * Click **Deploy** and copy the **Web App URL**.

---

## ⚙️ Customization Guide

This system can be tailored for other DAP offices or departments. Here is what to modify:

### 1. Changing Office Name & Logos

* **Files:** `Dashboard.html`, `Index.html`, `PrintSlip.html`, `Receipt.html`, etc.
* **Action:**
  * Search for `<div class="header">`.
  * Update the `<h2>` text (e.g., change "Finance Department" to "HR Department").
  * Replace the `<img>` src URL with your department's logo hosted on a public link (or Base64).

### 2. Modifying Document Types

* **File:** Google Sheet (`DocTypes` tab).
* **Action:** Simply add, remove, or edit the list in Column A. The web app pulls this list dynamically; no code change is required.

### 3. Customizing Email Templates

* **Files:** `Code.gs` and `InternalLogging.js`.
* **Functions:** Look for functions starting with `send...` (e.g., `sendLiaisonReceiptEmail`, `sendRoutedEmail`).
* **Action:** You can edit the HTML strings within these functions to change the email wording, color scheme (currently DAP Blue `#1C2790`), or layout.

---

## 🔒 Security & Privacy Note

* **Initials:** The system collects digital initials for verification purposes only. These images are stored securely in the designated Google Drive folder.
* **Data Access:** Only users with access to the backing Google Sheet can view the raw database.
* **Scope:** The script runs with the permissions of the account that deployed it (`Execute as: Me`).

---

## 📜 License

* **Copyright:** Development Academy of the Philippines (DAP)
* **Author:** Finance Department - Office of the Department Manager
