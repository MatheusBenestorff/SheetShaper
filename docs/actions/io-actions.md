# IO Actions: Load & Save

These actions are the entry and exit points of any SheetShaper pipeline. They handle the transfer of data between the file system (Disk) and the engine's memory.

## 1. LoadSource

The **LoadSource** action reads an Excel file (`.xlsx`) from the mapped `input/` folder and loads it into memory. You must assign an **Alias** to the loaded workbook, which acts as a variable name for subsequent steps.

### Configuration

```json
{
  "stepId": "1",
  "action": "LoadSource",
  "params": {
    "file": "sales_data_2024.xlsx",
    "alias": "RawData"
  }
}
```

## Parameters

| Parameter | Type   | Required | Description                                                                                                 |
| :-------- | :----- | :------- | :---------------------------------------------------------------------------------------------------------- |
| file      | String | Yes      | The filename (including extension) located in the input/ folder.                                            |
| alias     | String | Yes      | A unique internal name (ID) to reference this workbook in later steps (e.g., "Source", "Data", "Template"). |

## 2. SaveFile

The **SaveFile** action takes a workbook currently held in memory and writes it to the mapped output/ folder.

### Configuration

**Option A:** Explicit Source (Recommended for complex pipelines) Specifies exactly which workbook (Alias) to save.

```json
{
  "stepId": "5",
  "action": "SaveFile",
  "params": {
    "fileName": "final_report.xlsx",
    "sourceAlias": "CleanData"
  }
}
```

**Option B:** Implicit Source (Simple pipelines) If sourceAlias is omitted, the engine will save the first available workbook found in memory.

```json
{
  "stepId": "2",
  "action": "SaveFile",
  "params": {
    "fileName": "copy.xlsx"
  }
}
```

## Parameters

| Parameter   | Type   | Required | Description                                                                                             |
| :---------- | :----- | :------- | :------------------------------------------------------------------------------------------------------ |
| fileName    | String | Yes      | The name of the file to be created in the output/ folder.                                               |
| sourceAlias | String | Yes      | The Alias of the workbook you want to save. If not provided, the engine saves the first workbook found. |

## Best Practice

In pipelines with multiple steps (e.g., Load -> Filter -> Pivot -> Map), you often have multiple versions of data in memory ("Raw", "Filtered", "Pivoted"). Always use sourceAlias to ensure you are saving the correct version of your data.
