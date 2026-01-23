using ClosedXML.Excel;

namespace SheetShaper.Core.Actions;

public class AddFormulaColumnAction : ISheetAction
{
    public string ActionName => "AddFormulaColumn";

    public void Execute(ActionContext context)
    {
        
        string alias = context.GetParam("sheet");
        string headerName = context.GetParam("newHeader");
        string formulaTemplate = context.GetParam("formula"); // Ex: "A{row} + B{row}"
        
        string? format = context.GetOptionalParam("format");  // Ex: "dd/MM/yyyy" or Currency

        if (!context.Workbooks.ContainsKey(alias))
            throw new Exception($"Workbook '{alias}' not loaded.");

        var wb = context.Workbooks[alias];
        var ws = wb.Worksheets.First(); 

        Console.WriteLine($"     Adding formula column '{headerName}' to '{alias}'...");

        var lastColUsed = ws.LastColumnUsed();
        int newColIndex = (lastColUsed?.ColumnNumber() ?? 0) + 1;
        var newColumn = ws.Column(newColIndex);

        ws.Cell(1, newColIndex).Value = headerName;
        ws.Cell(1, newColIndex).Style.Font.Bold = true;

        // Applies the formula line by line
        var rows = ws.RangeUsed().RowsUsed().Skip(1);

        foreach (var row in rows)
        {
            int rowNum = row.RowNumber();
            
            // Replace the {row} placeholder with the actual number (e.g., 2, 3, 4...)
            string finalFormula = formulaTemplate.Replace("{row}", rowNum.ToString());
            
            // Sets the cell in the new column
            var cell = ws.Cell(rowNum, newColIndex);
            
            cell.FormulaA1 = finalFormula;

            if (!string.IsNullOrEmpty(format))
            {
                cell.Style.DateFormat.Format = format; 
                cell.Style.NumberFormat.Format = format;
            }
        }
        
        newColumn.AdjustToContents();
    }
}