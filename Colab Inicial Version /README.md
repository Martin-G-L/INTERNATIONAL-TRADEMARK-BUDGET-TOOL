# 🧾 Trademark Budget Generator

A robust Python solution for automating the generation of **budget documents** for trademark **applications** ("Depósito") and **renewals** ("Prorrogação"). This tool reads data from structured Excel spreadsheets and produces professional-grade **Word (.docx)** and **PDF** documents using dynamic LaTeX templates. Built to support single or multi-trademark workflows, this tool is optimized for deployment in **Google Colab**, with integrated **Google Drive** storage and Brazilian Portuguese formatting.

---

## 📌 Overview

This script facilitates:

- 💼 Budgeting for **trademark registration** and **renewal** processes  
- 📊 Reading and parsing data from Excel (`.xlsx`)
- 📝 Dynamic generation of `.docx` files using pre-configured templates
- 📄 LaTeX-based compilation of **PDF** budget summaries and correspondent details

---

## 🚀 Features

- **🔁 Input Flexibility**  
  Budget for single or multiple trademarks, with support for multiple countries and correspondents.

- **📄 Document Generation**  
  Creates DOCX and PDF files using structured Word/LaTeX templates for both budget and procedural details.

- **📥 Data Handling**  
  Reads Excel data, normalizes columns, and sanitizes input for LaTeX compatibility.

- **🌎 Localization**  
  Brazilian Portuguese formatting for dates, currency, and special characters.

- **🛠️ Error Handling**  
  Input validation, debugging logs, and failsafe handling of missing files.

- **🎨 Template System**  
  Modular and customizable LaTeX and Word templates for consistent branding and formatting.

---

## 🧰 Prerequisites

- Python ≥ 3.7  
- Google Colab (recommended environment)  
- Google Drive access (input/output storage)  
- LaTeX with `XeLaTeX` support (auto-installed in Colab)  
- Required Python libraries (see below)

## 🗂️ Input Preparation

Place the following files in your **Google Drive**:

- **Excel File:**  
  

- **Word Templates:**  


- **Logo File:**  


---

## 🧪 Usage

Execute the notebook in **Google Colab** and follow the interactive prompts:

1. Input **company name (Titular)**
2. Choose **single** or **multi-trademark** budgeting
3. Specify **country and correspondent** (if applicable)
4. Provide trademark details:
   - Name
   - Type: `Mista`, `Nominativa`, or `Figurativa`
   - Number of classes
   - Choose between `Depósito` or `Prorrogação Ordinário / Extra Ordinário`
5. To exit at any prompt: type `SAIR`

---

## 📤 Output

### 📄 Word Documents  
`.docx` budgets are saved to:  
`/content/drive/MyDrive/Beerre/Projeto Schedule of Fees/Orcamentos/`

### 📄 PDF Documents  
- Combined **Budget PDF**
- **Correspondent Details** PDF

### 🔖 File Naming Convention  
- `Orçamento_<Trademark>_<Type/Country>_<Date>.docx/pdf`  
- `Detalhes_<Trademark>_<Country>_<Date>.pdf`

---

## 🗃️ File Structure

```bash
trademark-budget-generator/
├── 📁 STEP 1: Set Up Libraries and Read Excel
│   ├── main.py                         # Imports pandas, python-docx, etc.
│   ├── PLANILHAS/
│   │   └── trademark_data.xlsx         # Input Excel with trademark data

├── 📁 STEP 2: Prepare Regional Formatting of Dates and Characters
│   └── main.py                         # Localized date, currency, and LaTeX text sanitization

├── 📁 STEP 3: Prepare LaTeX Templates
│   ├── main.py                         # LaTeX templates writen directly in the code



├── 📁 STEP 4: Prepare LaTeX and DOCX Generating Functions
│   └── main.py                         # Functions for document rendering (.docx/.pdf)

├── 📁 STEP 5: Gather User Inputs
│   └── main.py                         # Prompts for Titular, type, classes, country, etc.

├── 📁 STEP 6: Generate Corresponding Files
│   └── main.py                         # Dynamic generation of DOCX and LaTeX per trademark

├── 📁 STEP 7: Save Corresponding Files
│   └── Orcamentos/
│       ├── Budget_<...>.docx           # Word output
│       ├── Budget_<...>.pdf            # LaTeX-generated PDF
│       └── Details_<...>.pdf           # Correspondent-specific information PDF

├── README.md                           # Documentation


## 📦 Dependencies

### Python Libraries
Ensure the following libraries are installed in your environment:

- `pandas`
- `unicodedata`
- `tabulate`
- `gdown`
- `python-docx`
- `pytz`
- `re`
- `subprocess`
- `math`

### LaTeX Packages
Required LaTeX packages for PDF generation via XeLaTeX:

- `texlive-xetex`
- `texlive-fonts-recommended`
- `texlive-latex-recommended`
- `texlive-lang-portuguese`

---

## ⚠️ Notes

### Colab-Optimized
This project is optimized for **Google Colab**, leveraging:

- Integrated Google Drive access for file I/O
- Pre-installed LaTeX environment for XeLaTeX PDF generation

> 🛠️ **Note:** If running locally, you must manually install and configure LaTeX, and update file paths accordingly.

### LaTeX Compilation
- PDF generation relies on **XeLaTeX**.
- Compilation is executed **twice** to resolve all references and formatting.

### Debugging
- Built-in debug logs assist with:
  - Verifying column names from the Excel file
  - Validating LaTeX placeholder values
  - Tracing execution flow and subprocess output

### Assumptions
- The script expects:
  - A correctly formatted Excel input
  - Predefined Word and LaTeX templates placed in expected directories
  - Filenames and paths must not be altered unless adapted in code

### Customization
- Update the `honorarios_dict` dictionary to reflect fee changes
- Modify Word/LaTeX templates to suit different formatting or document structures

---

## 🤝 Contributing

We welcome and encourage contributions from the community. To contribute:

1. Fork the repository
2. Create a new branch for your feature or fix
3. Submit a Pull Request (PR) for review

### ✅ Contribution Guidelines

- Ensure your code follows **PEP 8** standards
- Include documentation for any new features or modules
- Maintain compatibility with existing **XeLaTeX** templates and environment

---
## 📬 Contact

For questions, suggestions, or collaboration inquiries, please feel free to reach out:

- **👤 Author**: Martin G. Lartigue  
- **📧 Email**: [martin.g.lartigue@gmail.com](mailto:martin.g.lartigue@gmail.com)

I'm open to feedback, feature requests, and contributions related to this project or other automation initiatives in the IP (Intellectual Property) and document generation space.
