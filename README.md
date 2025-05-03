# Vulnerability Report Generator 🛡️📄

**Automate penetration testing reports in multiple formats with a single interface.**  
Save hours on vulnerability documentation by generating polished PDF, Word, Excel, and HTML reports instantly.

[![Python Version](https://img.shields.io/badge/Python-3.8%2B-blue)](https://www.python.org/)
[![License: MIT](https://img.shields.io/badge/License-MIT-green)](https://opensource.org/licenses/MIT)


## Features ✨
- **Multi-Format Export**: Generate reports in PDF, Word (DOCX), Excel (XLSX), and HTML
- **GUI Interface**: User-friendly Tkinter-based interface
- **Vulnerability Templates**: Pre-defined sections (Description, Impact, PoC, etc.)
- **Severity Highlighting**: Color-coded Critical/High/Medium/Low ratings
- **Smart Defaults**: Auto-"N/A" for empty fields
- **Local Execution**: No cloud dependency for sensitive data

## Installation ⚙️

### Requirements
- Python 3.8+
- Tkinter (usually included with Python)

### Steps
1. Clone the repository:
   ```bash
   git clone https://github.com/nihaltikka/vulnerability-report-generator.git
   cd vulnerability-report-generator
2. Install dependencies:
   ```bash
   pip install -r requirements.txt

   # Linux users: Install Tkinter separately:
     ```bash
     sudo apt-get install python3-tk
### Usage 🖥️
1. Start the application:
   ```bash
   python3 report_generator.py
3. Workflow

  Enter Report Name

  Fill vulnerability details:
  
  Vulnerability Name (required)
  
  Description, Impact, Affected URL, etc.
  
  Select Severity from dropdown
  
  Click "➕ Add Vulnerability"
  
  Repeat for multiple findings
  
  Choose output format (PDF/Word/Excel/HTML)
  
  Click "🛠 Generate Report"
  
3. Find outputs in the project directory with timestamps:

ReportName_2023-10-25.pdf/docx/xlsx/html

### Technical Details 🔧
    Built With:
    
    GUI: Tkinter
    
    PDF: ReportLab + FPDF2
    
    Word: python-docx
    
    Excel: pandas + XlsxWriter
    
    HTML: Native templating

### Project Structure:

  .
├── report_generator.py    # Main application
├── requirements.txt       # Dependencies
├── sample_reports/        # Example outputs
└── README.md              # This file

### Contributing 🤝

Found a bug? Want a feature?

1. Fork the repo

2. Create your branch: git checkout -b feature/your-feature

3. Commit changes: git commit -m 'Add amazing feature'

4. Push: git push origin feature/your-feature

5. Open a PR!


Disclaimer: This tool is for educational/authorized testing purposes only.
Need Help? Open an issue or reach out via LinkedIn(https://www.linkedin.com/in/nihaltikka)
Read me on: https://elcazad0r.medium.com/
Portfolio : https://nihaltikka.github.io/
