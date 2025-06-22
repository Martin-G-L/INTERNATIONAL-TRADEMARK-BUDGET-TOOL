# 🛠 Trademark Budget Tool – Python Version

This folder contains the **standalone Python (`.py`) version** of the **Trademark Budget Tool** developed by M-G-L International.
It's a extension with adapted path to the colab version contained in the Repository. See Colab Inicial Version README.md for more details.

## 📌 Description

The tool facilitates the budgeting process for international trademark services such as:
- **Application (Depósito)**
- **Ordinary Term Extension (Prorrogação Ordinária)**
- **Extraordinary Term Extension (Prorrogação Extraordinária)**

It provides a user-friendly graphical interface built with **Tkinter**, ensuring better usability for legal and administrative workflows.

## ✅ Key Enhancements

- **PDF Generation Support (XeLaTeX):** Adjustments made to ensure compatibility with `xelatex` for compiling structured budget documents.
- **Improved UX Design:** Refined navigation and user interface components using `Tkinter`'s `ttk` styling features.
- **Robust Logging:** Logs are now organized in a dedicated directory with improved clarity and structure for debugging or audit purposes.

## 🚀 Getting Started

To run this version locally:

1. Ensure you have the following installed:
   - Python 3.x
   - Tkinter (included by default with most Python distributions)
   - XeLaTeX (must be installed and accessible via system PATH)

2. **Configure the required file paths** in the script before execution:
   - Paths for the logo image, PDF output files, logs, Excel spreadsheet with correspondent data, and Word templates must be set according to your system.

3. You have two options:
   - ✅ **Customize paths manually** to suit your project directory structure.
   - 🔁 **Alternatively**, if you'd like to use the tool with the original default paths, place the required assets in the following locations:
  
     
C:\Users\Your Username\Documents\PROGRAM TESTING FOLDERS\exes\Project Schedule of Fees\CODE\OUTPUT\LOGS
→ **Log files directory**

C:\Users\Your Username\Documents\PROGRAM TESTING FOLDERS\exes\Project Schedule of Fees\CODE\OUTPUT
→ **PDF output files directory**

C:\Users\Your Username\Documents\PROGRAM TESTING FOLDERS\exes\Project Schedule of Fees\CODE\PLANILHAS
→ **Excel (.xlsx) and Word (.docx) templates directory**

C:\Users\Your Username\Documents\PROGRAM TESTING FOLDERS\exes\Project Schedule of Fees\CODE\Logo
→ **Logo image directory**

