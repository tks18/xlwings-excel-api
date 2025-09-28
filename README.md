# 📊 Xlwings Excel Custom API

A modular **Python–Excel–VBA automation framework** for building and deploying **User Defined Functions (UDFs)**, query helpers, and reusable automation workflows.
This project centralizes text, table, and data transformation utilities across **Excel**, **Power Query**, and **Power BI**, making it easier to maintain, test, and extend automation logic.

---

## 🚀 Features

* **Excel–Python UDFs**

  * Text manipulation (slugify, regex search/replace, string cleaning).
  * Table automation (dynamic CSV/PDF loaders, column processors).
  * Data validation & formatting helpers.

* **Power Query Function Library**

  * Pre-packaged **M functions** with documentation metadata.
  * Organized into categories (Text, Table, Regex, Loaders, etc.).
  * Auto-folder structure inside Excel (`_fx_queries`).

* **Integration Workflows**

  * VBA macros to register & call Python UDFs.
  * Outlook/automation integrations for reporting.
  * Power Automate safe testing patterns.

---

## ⚙️ Installation & Setup

1. **Clone the repository**

   ```bash
   git clone https://github.com/tks18/xlwings_excel_api.git
   cd xlwings_excel_api
   ```

2. **Install Python dependencies**

   ```bash
   uv sync
   ```

3. **Enable the custom add-in in Excel**

   * Open Excel → File → Options → Add-ins → Excel Addins → Enable.
   * Go to Shan's Labs → Import Functions.
   * Import the Power Query Function using the search bar in the same tab.

---

## 🛠 Usage

### Example: Slugify a String (Excel UDF)

```excel
=SLUG_BASIC("Hello World!")
```

→ `hello-world`

### Example: Load CSV with Dynamic Columns (Power Query)

```m
fx_LoadCSV("C:\data\sales.csv", 200)
```

---

## 📖 Documentation

* Each UDF/function includes **YAML front-matter** for:

  * `name` – function name
  * `category` – grouping for folders
  * `tags` – keywords
  * `description` – usage notes
  * `version` – function version

Example:

```yaml
---
name: fx_LoadPDFDynamic
category: Table Loaders
tags: [pdf, dynamic, import, columns, table]
description: "Load all pages of a PDF into a flattened table by expanding dynamic columns."
version: "v2.1"
---
```

---

## 🤝 Contributing

1. Fork the repo.
2. Create a feature branch (`api/my-new-func`).
3. Commit changes with descriptive messages.
4. Submit a pull request.

---