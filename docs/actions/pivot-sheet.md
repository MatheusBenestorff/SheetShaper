# Action: PivotSheet

The **PivotSheet** action transforms data from a "Long Format" (Entity-Attribute-Value) to a "Wide Format" (Tabular). This is commonly used when handling exports from ERP systems (like SAP) where attributes of a single entity are split across multiple rows.

## Concept

It takes a list where characteristics are rows and turns them into columns, grouping by a unique identifier.

**Before (Source - Long Format):**
| Row | Key (Col A) | Value (Col B) | ... | ID (Col E) |
| :-- | :--- | :--- | :-- | :--- |
| 1   | Weight      | 10kg          | ... | OV-100     |
| 2   | Color       | Blue          | ... | OV-100     |
| 3   | Color       | Green         | ... | OV-200     |

**After (Target - Wide Format):**
| ID | Color | Weight |
| :--- | :--- | :--- |
| OV-100  | Blue  | 10kg   |
| OV-200  | Green |        |

## Configuration

```json
{
  "stepId": "2",
  "action": "PivotSheet",
  "params": {
    "sourceSheet": "Source",
    "targetSheet": "Report",
    "groupByColumn": "E",
    "pivotKeyColumn": "A",
    "pivotValueColumn": "B"
  }
}
```

## Parameters

|  Parameter  | Type | Required | Description |
| :---------- | :--- | :------- | :---------- |
| sourceSheet | String | Yes | The alias of the workbook containing the raw data. |
| targetSheet | String | Yes | The alias for the new workbook that will be created with the pivoted data. |
| groupByColumn | String | Yes | The column letter (e.g., "E") containing the unique ID of the entity (the Primary Key). |
| pivotKeyColumn | String | Yes | The column letter (e.g., "A") containing the attribute names. These will become the Headers of the new columns. |
| pivotValueColumn | String | Yes | The column letter (e.g., "B") containing the actual values. |

## Real World Use Case: SAP Order Consolidation

**Problem:** An SAP export contains 50,000 rows. Each sales order (OV) has multiple lines, one for each characteristic (Weight, Size, Color, Shipping Method). Goal: Create a report with one line per Sales Order.

**Solution:** Use PivotSheet grouping by the Order Number column. The engine will automatically detect all unique characteristics (keys) and create the necessary columns dynamically.