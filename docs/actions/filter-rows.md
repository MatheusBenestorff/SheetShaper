# Action: FilterRows

The **FilterRows** action allows you to exclude rows from a dataset based on specific conditions. It creates a new workbook containing only the rows that match the criteria.

## Configuration

```json
{
  "stepId": "2",
  "action": "FilterRows",
  "params": {
    "sourceSheet": "RawData",
    "targetSheet": "CleanData",
    "column": "C",
    "operator": "GreaterThan",
    "value": 1000
  }
}
```

## Parameters

|  Parameter  | Type | Required | Description |
| :---------- | :--- | :------- | :---------- |
| sourceSheet | String | Yes | Alias of the input workbook. |
| targetSheet | String | Yes | Alias for the new filtered workbook. |
| column | String | Yes | The column letter to evaluate (e.g., "C"). |
| operator | String | Yes | The comparison rule (see list below). |
| value | Mixed | Yes | The reference value. Can be a Number (1000) or String ("Approved"). |


## Supported Operators

|  Operator | Description |
| :---------| :---------- |
| Equals | Exact match (Case insensitive for text). |
| NotEquals | Everything except the value. |
| GreaterThan | Numeric comparison. |
| LessThan | Numeric comparison. |
| Contains | Checks if the text contains the value (substring). |

## Example

To keep only rows where the Status (Column D) is "Active": 

- "column": "D"
- "operator": "Equals"
- "value": "Active"
