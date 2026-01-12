# Action: MapColumns

The **MapColumns** action is used to select, rename, and reorder columns. It creates a clean dataset containing only the specified columns from the original file.

## Configuration

```json
{
  "stepId": "2",
  "action": "MapColumns",
  "params": {
    "sourceSheet": "Source",
    "targetSheet": "FinalReport",
    "mappings": [
      { "from": "A", "to": "A", "header": "Transaction Date" },
      { "from": "F", "to": "B", "header": "Total Amount" }
    ]
  }
}
```

## Parameters

|  Parameter  | Type | Required | Description |
| :---------- | :--- | :------- | :---------- |
| sourceSheet | String | Yes | Alias of the input workbook. |
| targetSheet | String | Yes | Alias for the new filtered workbook. |
| mappings | Array | Yes | A list of mapping objects defining the transformation. |

## Mapping Object Structure

| Key | Description |
| :---------| :---------- |
| from | The source column letter (e.g., "F"). |
| to | The destination column letter (e.g., "B"). |
| header | The new name for the column header (Row 1). |

## Use Case

You receive a spreadsheet with 50 columns, but your system only needs the Date and the Total. Use MapColumns to extract just these two, placing them in columns A and B, and renaming headers to English.