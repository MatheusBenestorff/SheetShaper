# Action: AddFormulaColumn

The **AddFormulaColumn** action allows you to enrich your dataset by creating new columns based on Excel formulas. instead of hardcoding mathematical logic in the ETL engine, you leverage the native power of Excel calculation.

This is useful for mathematical operations (`Price * Qty`), text manipulation (`Concatenate`), or logical conditions (`IF`).

## Concept

It appends a new column to the right of the existing data and injects a dynamic formula into every row.

**The `{row}` Placeholder:**
The engine uses a special token `{row}` which is automatically replaced by the current row number during execution.
* Formula Template: `A{row} * B{row}`
* Result in Row 2: `=A2 * B2`
* Result in Row 3: `=A3 * B3`

**Before:**
| Qty (Col A) | Price (Col B) |
| :--- | :--- |
| 10 | 5.00 |
| 2  | 20.00 |

**After (Formula: `A{row}*B{row}`):**
| Qty | Price | Total |
| :--- | :--- | :--- |
| 10 | 5.00 | **50.00** |
| 2  | 20.00 | **40.00** |

## Configuration

```json
{
  "stepId": "4",
  "action": "AddFormulaColumn",
  "params": {
    "sheet": "SalesData",
    "newHeader": "Total",
    "formula": "A{row} * B{row}",
    "format": "$ #,##0.00" 
  }
}
```

## Parameters

|  Parameter  | Type | Required | Description |
| :---------- | :--- | :------- | :---------- |
| sheet | String | Yes | Alias of the target workbook. |
| newHeader | String | Yes | The name of the header for the newly created column. |
| formula | String | Yes | The Excel formula template. Must use English function names (e.g., SUM, IF) and the {row} placeholder for relative references. |
| format | String | No | An optional Excel format string (e.g., "dd/MM/yyyy" for dates or "$ #,##0.00" for currency). |

## Common Patterns

1. **Mathematical Calculation**

Calculate total price including tax (assuming Tax Rate is in Col C).

- Formula: `(A{row} * B{row}) * (1 + C{row})`

2. **Text Concatenation**

Join First Name (Col A) and Last Name (Col B).

- Formula: `A{row} & " " & B{row}`

Or: CONCATENATE(A{row}, " ", B{row})

3. **Logical Conditions `(IF)`**

Flag high-value orders.

- Formula: `IF(C{row} > 1000, "High Priority", "Standard")`

## Important Notes

**Language:** Formuls must be written using English syntax (e.g., use IF instead of SE, SUM instead of SOMA), regardless of your PC's local settings.

**Separators:** Use commas (,) to separate arguments in functions, not semicolons (;), as per standard US/International Excel syntax