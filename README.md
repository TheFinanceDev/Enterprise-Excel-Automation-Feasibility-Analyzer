# 🚀 Enterprise Excel Automation Analyzer

A robust and secure Python tool for assessing the automation feasibility of Excel workbooks in enterprise environments. It delivers in-depth analysis of spreadsheet structure, formula complexity, and business process patterns — empowering organizations to make data-driven decisions on automation strategies and tool selection. 💼📊🔒

---

## 📋 Table of Contents

- [About The Project](#about-the-project)  
- [Features](#⚙️-features)  
- [Getting Started](#🚀-getting-started)  
- [Usage](#🏃-how-to-use)  
- [Requirements](#📦-requirements)  
- [Contributing](#🤝-contributing)  
- [License](#📜-license)  
- [Contact](#📫-contact) 

---

## About The Project

Managing and automating Excel files in large organizations can be complex due to varied file structures, formula intricacies, and process workflows. This tool helps:

- **Analyze** key aspects of Excel workbooks that impact automation feasibility  
- **Score** structure, formula complexity, and pattern recognition  
- **Identify** risks and blockers like hidden sheets and macros  
- **Recommend** appropriate automation tools including Python, VBA, Power Platform, and RPA  
- **Generate** clear reports for stakeholders  

Ideal for business analysts, process owners, and automation engineers aiming to streamline Excel-based workflows securely and efficiently.

---

## ⚙️ Features

- ✔ Validates Excel files: formats, access, size  
- ✔ Detailed sheet and workbook structure analysis  
- ✔ Assesses formula complexity with categorization  
- ✔ Detects automation-friendly patterns (templates, time-based reports, consolidation)  
- ✔ Flags risks like macros, protection, and data inconsistencies  
- ✔ Provides automation feasibility score and effort estimation  
- ✔ Recommends suitable automation tools and strategies  
- ✔ Interactive CLI and report exporting capabilities  
- ✔ Entirely local processing with no external data transmission  

---

## 🚀 Getting Started

### Prerequisites

Make sure you have Python 3.7 or newer installed. Then install required packages:

```
pip install pandas openpyxl
```

Alternatively, if you provide a `requirements.txt`, install using:

```
pip install -r requirements.txt
```


---

## 🏃 How To Use

### Interactive Mode

Run the tool without arguments to start interactive CLI:

```
python excel_automation_analyzer.py
```


You will be shown a menu where you can:

- Analyze Excel files by entering their full path
- Export reports
- View system info

Follow on-screen prompts.

### Command Line Mode

Analyze a specific Excel file directly by providing the path as an argument:

```
python excel_automation_analyzer.py path/to/your_excel_file.xlsx
```


Optional flags:

- `--quiet` : Minimize console output  
- `--no-export` : Skip report export  

---

## 📦 Requirements

- Python 3.7+  
- pandas  
- openpyxl  

Install missing packages with `pip`.

---

## 🤝 Contributing

Contributions are welcome! To contribute:

1. Fork the repository  
2. Create a feature branch (`git checkout -b feature/NewFeature`)  
3. Commit your changes (`git commit -m 'Add NewFeature'`)  
4. Push to the branch (`git push origin feature/NewFeature`)  
5. Open a pull request  

Please follow the code style and keep commits clear.

---

## 📜 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

---

## 📫 Contact

Enterprise Analytics Team  
Email: analytics-team@example.com  
GitHub: [https://github.com/yourusername](https://github.com/yourusername)  

---

✨ _Thank you for checking out this project!_
