# SheetShaper

![Build Status](https://github.com/MatheusBenestorff/SheetShaper/actions/workflows/dotnet.yml/badge.svg)

SheetShaper empowers developers and data analysts to automate Excel tasks (like merging columns, cleaning data, and reformatting) without writing ad-hoc scripts. You define a **Pipeline** of steps in a JSON file, and the engine executes the transformations sequentially.


## Project Structure

To use SheetShaper, organize your files within the mapped volume (default: `examples/`):

```text
examples/
├── config/          <-- Place your pipeline definition (JSON)
│   └── sales_cleanup.json
├── input/           <-- Place source Excel files
│   └── raw_data_2024.xlsx
└── output/          <-- The engine saves the processed files here
    └── clean_report_2024.xlsx
```

## Configuration Guide

The core concept of SheetShaper is the Pipeline. A job consists of a list of steps executed in order.

**Pipeline Structure**

Each step requires an action and a set of params.

```json
{
  "jobName": "Sales_Data_Cleanup",
  "steps": [
    {
      "stepId": "1",
      "action": "LoadSource",
      "params": {
        "file": "teste.xlsx",
        "alias": "Source"
      }
    },
    {
      "stepId": "2",
      "action": "MapColumns",
      "params": {
        "sourceSheet": "Source",
        "targetSheet": "Report",
        "mappings": [
          { "from": "A", "to": "A", "header": "Data Venda" },
          { "from": "F", "to": "B", "header": "Receita Total" }
        ]
      }
    },
    {
      "stepId": "3",
      "action": "SaveFile",
      "params": {
        "fileName": "clean_report_2024.xlsx",
        "sourceAlias": "Report"
      }
    }
  ]
}
```

## Documentation

Detailed guide for available Actions:

* [Load & Save](actions/io-actions.md)
* [Map Columns](actions/map-columns.md)
* [Filter Rows](actions/filter-rows.md)
* [Pivot Sheet](actions/pivot-sheet.md)

### Architecture

- **SheetShaper.Core:** The pipeline execution engine using ClosedXML.
- **SheetShaper.CLI:** The command-line runner that handles File I/O and configuration loading.

### Quick Start (Docker)

You can run SheetShaper using Docker Compose without installing the .NET SDK.

1.  **Setup:** Place your excel file in examples/input and your JSON in examples/config

2.  **Run the engine:**

    ```bash
    docker compose up
    ```



