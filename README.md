# docx-mapper

This is a simple script that takes a Word document template and an Excel file with values to replace placeholders, and generates a finalized document with the data filled in.

## 🧾 What it does

- Reads `base_document.docx`, which contains placeholders like `{{company}}`, `{{first_name}}`, etc.
- Loads `mapping.xlsx`, which contains the field names and the values to replace.
- Outputs a new file `final_document.docx` with all placeholders replaced.

## 🐳 Prerequisites

- [Docker Desktop](https://www.docker.com/products/docker-desktop/)
- Docker Compose

## 🗂️ Folder structure

Here is the folder structure for the project:

```
project_root/
├── backend/                # Contains the main application files
│   ├── base_document.docx  # Word template with placeholders
│   ├── mapping.xlsx        # Excel file with placeholder values
│   ├── script.py           # Main script for processing
│   ├── Dockerfile          # Docker configuration for the backend
│   ├── requirements.txt    # Python dependencies
├── LICENSE                 # License file for the project
├── README.md               # Documentation for the project
└── docker-compose.yml      # Docker Compose configuration
```

## 🛠️ How to use

1. Place your Word template (`base_document.docx`) and Excel file (`mapping.xlsx`) in the root folder.
2. Open a terminal and run:
```sh
docker compose up --build
```
3. The output file final_document.docx will be generated in the same folder.

## ⚙️ Optional settings
The script accepts a flag called USE_DATE_FORMAT that determines how Excel date fields are formatted:
- True → formats dates as dd/mm/yyyy (e.g., 01/01/1980)
- False (default) → keeps the full datetime (e.g., 01/01/1980 00:00:00)

To enable this option, edit the script.py and set:
```py
USE_DATE_FORMAT = False
```

## 🧪 Example
Given this placeholder in the Word file:

```doc
{{name}} was born on {{born_in}}
```
And the following Excel content:
```xls
A             B
name          Mario Rossi
born_in       01/01/1980
```
You will get this result:

```doc
Mario Rossi was born on 01/01/1980
```
