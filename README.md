# Excel Data Quality Validator

## Overview

A generic tool designed to validate the **structure and data quality of Excel files (`.xlsx`)** before they are used in data pipelines.

The goal is to provide an **early-stage data observability layer**, automatically detecting structural issues and data anomalies without requiring predefined rules per file.

---

## Features

### Structural Validation

* Detects added or removed sheets
* Detects added or removed columns
* Validates column order

### Schema Validation

* Automatic data type detection per column
* Detects type changes

### Data Quality Checks

* Null percentage per column
* Outlier detection (IQR method)
* Descriptive statistics:

  * Mean
  * Standard deviation
  * Min / Max
  * Distinct count

### Drift Detection

* Mean variation
* Standard deviation variation
* Cardinality (distinct) changes

### Smart Alerts

* Severity classification:

  * CRITICAL
  * WARNING
  * INFO

### Quality Score

* Score from 0 to 100 based on detected issues

---

## Concept

The tool works with a **baseline approach**:

1. First run → generates a baseline (`JSON`)
2. Next runs → compare current file against baseline
3. Logs any detected differences

---

## Architecture

```text id="x2c7bt"
Source → Load → Analyze → Compare → Log
```

---

## Supported Sources

* 📁 Local files
* ☁️ SharePoint (via OneDrive sync)
* 🧊 Azure Blob Storage

---

## Configuration

All parameters are defined at the top of the script:

```python id="f3g2je"
ORIGEM_ARQUIVO = 'local'  # local | sharepoint | azure_blob
```

### Local / SharePoint

```python id="w3o2h9"
CAMINHO_ARQUIVO = r"C:\\folder\\"
NOME_ARQUIVO = "file.xlsx"
```

### Azure Blob

```python id="oqx4rm"
AZURE_CONNECTION_STRING = ""
AZURE_CONTAINER_NAME = ""
AZURE_BLOB_NAME = ""
```

---

## Usage

### 1. Install dependencies

```bash id="k8p2ls"
pip install pandas openpyxl azure-storage-blob
```

---

### 2. Generate baseline

```python id="8b1zqk"
MODO_BASELINE = "yes"
```

Run:

```bash id="u9k2qn"
python checkExcel.py
```

---

### 3. Validate file

```python id="s7d9pw"
MODO_BASELINE = "no"
```

Run again:

```bash id="x9f1ar"
python checkExcel.py
```

---

## Outputs

### Baseline (JSON)

* Full structure and metrics of the file

### Log file

* Detected differences
* Severity summary
* Quality score

---

## Example Output

```text id="j0p2rm"
Summary:
CRITICAL: 2
WARNING: 3
INFO: 1

Quality Score: 85
```

---

## Use Cases

* Pre-ingestion validation for data pipelines
* Data quality monitoring
* Drift detection in Excel datasets
* Data governance support

---

## Limitations

* Does not auto-correct data
* Does not replace human analysis
* SharePoint API integration not yet implemented

---

## Roadmap

* SharePoint API integration
* Batch processing (multiple files)
* JSON output for pipeline integration
* Integration with orchestration tools (Airflow, ADF)

---

## Author

Developed as a generic solution for Excel data validation in data engineering environments.

---

