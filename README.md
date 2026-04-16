<div align="center">
    
# OFFICE-MORPH <br>v1.2
**Automated Legacy Office Modernization** <br>
by Manuel Pollhammer (2026)
</div>

---

## 📝 Description
**Office-Morph** is an intelligent automation utility designed to seamlessly convert legacy Microsoft Office binary formats into modern XML standards. It streamlines the transition from older archives to current, accessible formats.

## 📦 Components
*   **Offic-Morph.bat**: The interactive main menu and execution interface.
*   **FolderConverter.ps1**: The high-performance core processing engine.

## 🚀 Usage Modes
The tool is highly flexible and offers three distinct execution modes:

1.  **Drag-and-Drop (Maximum Convenience):** 
    Simply drag a folder and drop it directly onto the `Office-Morph.bat` file.
2.  **Manual Input:** 
    Launch the batch file and paste the target directory path into the console, then confirm with **Enter**.
3.  **Express Mode (Current Folder):** 
    Launch the batch file and press **Enter** at the path prompt to process the directory where the tool is located.

---

## 🛠️ New in v1.2: Maintenance Module
After converting your files, you can now use **Option [2]** in the main menu to:
*   **Deep Clean:** Recursively scan and permanently delete old `.doc`, `.xls`, and `.ppt` files.
*   **Safety First:** Includes a confirmation prompt to prevent accidental data loss.

---

## ⚠️ Important: Administrative Rights
This tool requires **LOCAL ADMINISTRATOR PRIVILEGES** to access Office COM interfaces and perform file system operations.  
**Please run `Offic-Morph.bat` by RIGHT-CLICKING and selecting "RUN AS ADMINISTRATOR".**

---

## ✨ Key Features
*   **Summary Statistics:** Provides a detailed report (Converted / Skipped / Errors) after each run. **✨NEW✨**
*   **Deep Scan:** Automatically detects legacy files across all subdirectories.
*   **Smart Skip:** Efficiently skips files already converted to save time.
*   **Temp-File Shield:** Automatically ignores hidden Office temporary files (`~$`). **✨NEW✨**
*   **Clean Naming:** Advanced logic prevents double file extensions (e.g., no `..xlsx`).

## 📋 Prerequisites
*   Installed Microsoft Office Suite (Word, Excel, PowerPoint).
*   Windows PowerShell 5.1 or higher.
*   Local Administrator permissions.

---

## 📸 Screenshots

<p align="center">
  <img src="./screenshots/screenshot_001.png" alt="Office-Morph Interface" width="800">
  <br>
  <i>Interface and Execution</i>
</p>

<p align="center">
  <img src="./screenshots/screenshot_002.png" alt="Office-Morph Interface" width="800">
  <br>
  <i>Interface and Execution</i>
</p>

<p align="center">
  <img src="./screenshots/screenshot_003.png" alt="Conversion Success" width="800">
  <br>
  <i>Successful Conversion Process</i>
</p>

---

**Developed by Manuel Pollhammer | Release 2026**

