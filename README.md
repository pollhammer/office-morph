<div align="center">
  
![ASCII Logo](./logo/logo.svg)
 # Office-Morph <br>v1.4
 **.doc, .xls, .ppt ➔ .docx, .xlsx, .pptx**<br>
  by Manuel Pollhammer (2026)
</div>

---

## 🚀 What is Office-Morph?
**Office-Morph** is an intelligent automation utility designed to seamlessly convert legacy Microsoft Office binary formats into modern XML standards. It streamlines the transition from older archives to current, accessible formats using the native Office COM engine.

## 📦 Components
*   **Office-Morph.bat**: The interactive main menu and execution interface.
*   **FolderConverter.ps1**: The high-performance core processing engine with advanced logging.

## 📝 Usage Modes
The tool is highly flexible and offers three distinct execution modes:

1.  **Drag'n'Drop (Maximum Convenience):** 
    Simply drag a folder and drop it directly onto the `Office-Morph.bat` file.
2.  **Manual Input:** 
    Launch the batch file and paste the target directory path into the console.
3. **Express Mode (Current Folder):**
   Press **Enter** without a path to process the tool's current directory.

---

## 🛠️ New in v1.4: Professional Logging & Stability
The latest update focuses on enterprise-grade reliability and transparency:

*   **Detailed Logging:** Generates `office_morph_details.log` in the target folder, documenting every single file path and its status.
*   **Smart Error Analysis:** Specifically identifies and logs reasons for failure, such as **Path Too Long** (Windows 260-char limit) or file locks.
*   **Silent Mode & Stability:** Improved background processing with `ReleaseComObject` to ensure Office processes are fully terminated after execution.
*   **Safety Confirmations:** Enhanced "Delete" module with file listing and [Y/N] confirmation to prevent accidental data loss.

---

## ✨ Key Features
*   **Detailed Reports:** Full breakdown (Converted/Skipped/Errors) displayed in console and saved to disk. **✨NEW✨**
*   **Path Length Guard:** Explicitly detects and logs path length issues on network drives. **✨NEW✨**
*   **Deep Scan:** Automatically detects legacy files across all subdirectories.
*   **Smart Skip:** Skips files already converted (e.g., if a .docx already exists) to save time.
*   **Temp-File Shield:** Automatically ignores hidden Office temporary files (`~$`).
---

## 📋 Prerequisites
*   Installed Microsoft Office Suite (Word, Excel, PowerPoint).
*   Windows PowerShell 5.1 or higher.
*   **Execution Policy:** Set to `Bypass` (handled automatically by the .bat launcher).

---

## 📸 Screenshots

<p align="center">
  <img src="./screenshots/screenshot_v1.3.001.png" alt="Main Menu" width="800">
  <br>
  <i>Main Menu</i>
</p>

<p align="center">
  <img src="./screenshots/screenshot_v1.3.002.png" alt="Interface and Execution" width="800">
  <br>
  <i>Interface and Execution</i>
</p>

<p align="center">
  <img src="./screenshots/screenshot_v1.3.003.png" alt="Delete old (.doc, .xls, .ppt) files" width="800">
  <br>
  <i>Delete old (.doc, .xls, .ppt) files</i>
</p>

<p align="center">
  <img src="./screenshots/screenshot_v1.2.004.png" alt="Successful Conversion Process" width="800">
  <br>
  <i>Successful Conversion Process</i>
</p>

---
**Developed by Manuel Pollhammer | Release 2026**


