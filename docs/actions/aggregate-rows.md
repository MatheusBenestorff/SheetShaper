# Action: AggregateRows

The **AggregateRows** action allows you to summarize data by grouping rows based on a specific column and performing mathematical calculations (like Sum, Average, Count) on other columns.

This is the SheetShaper equivalent of an SQL `GROUP BY` statement or the "Values" section of an Excel Pivot Table.

## Concept

It transforms detailed transactional data into high-level summary reports.

**Before (Source - Transactions):**
| Region (Col C) | Sales Rep | Amount (Col F) |
| :--- | :--- | :--- |
| North | John | 100 |
| North | Doe  | 200 |
| South | Jane | 500 |

**After (Target - Aggregated by Region):**
| Group Key | Total Amount | Transaction Count |
| :--- | :--- | :--- |
| North | 300 | 2 |
| South | 500 | 1 |

## Configuration

```json
{
  "stepId": "2",
  "action": "AggregateRows",
  "params": {
    "sourceSheet": "CleanData",
    "targetSheet": "MonthlySummary",
    "groupByColumn": "C",  
    "operations": [
      { 
        "targetColumn": "Total Amount", 
        "sourceColumn": "F",            
        "type": "Sum"                   
      },
      { 
        "targetColumn": "Transaction Count", 
        "type": "Count"                      
      }
    ]
  }
}
```

## Parameters

|  Parameter  | Type | Required | Description |
| :---------- | :--- | :------- | :---------- |
| sourceSheet | String | Yes | Alias of the workbook containing the detailed data. |
| targetSheet | String | Yes | Alias for the new summary workbook. |
| groupByColumn | String | Yes | The column letter (e.g., "C") used to group the rows. |
| operations | Array | Yes | A list of calculation rules to apply to each group. |

## Operation Object Structure

Inside the **operations** list, each object defines a calculated column:

|  Key  | Required | Description |
| :---------- | :------- | :---------- |
| targetColumn | Yes | The name of the new header in the summary report. |
| type  | Yes | The calculation method (see table below). |
| sourceColumn  | Conditional | The column letter to calculate. Required for Sum, Average, Min, Max. Ignored for Count. |

## Supported Operation Types

|  Type | Description |
| :---------- | :---------- |
| Sum | Adds up all numeric values in the source column. |
| Average  | Calculates the mean value of the source column. |
| Count | Counts how many rows exist in the group. |
| Max | Finds the highest numeric value in the group. |
| Min | Finds the lowest numeric value in the group. |

## Real World Use Case: Sales Commission Report

Problem: You have a spreadsheet with **50,000** sales records. You need to pay commissions based on total sales per salesperson.

**Solution:**

- Use AggregateRows.

- Set groupByColumn to the Salesperson Name column.

- Add an operation type: "Sum" on the Total Value column.

- The result is a clean list with one row per salesperson and their total sales, ready for payroll.




