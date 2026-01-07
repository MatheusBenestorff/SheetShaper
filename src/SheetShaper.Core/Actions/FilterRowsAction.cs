using ClosedXML.Excel;

namespace SheetShaper.Core.Actions;

public class FilterRowsAction : ISheetAction
{
    public string ActionName => "FilterRows";

    public void Execute(ActionContext context)
    {
        string sourceAlias = context.GetParam("sourceSheet");
        string targetAlias = context.GetParam("targetSheet");
        string column = context.GetParam("column");
        string op = context.GetParam("operator");
        string targetValueRaw = context.GetParam("value"); 

        if (!context.Workbooks.ContainsKey(sourceAlias))
            throw new Exception($"Source workbook '{sourceAlias}' not loaded.");

        var sourceWb = context.Workbooks[sourceAlias];
        var sourceWs = sourceWb.Worksheets.First();

        var targetWb = new XLWorkbook();
        var targetWs = targetWb.Worksheets.Add("FilteredData");

        Console.WriteLine($"     Filtering '{sourceAlias}' where Column {column} {op} {targetValueRaw}...");

        var headerRow = sourceWs.Row(1);
        if (!headerRow.IsEmpty())
        {
            for (int i = 1; i <= headerRow.LastCellUsed().Address.ColumnNumber; i++)
            {
                targetWs.Cell(1, i).Value = headerRow.Cell(i).Value;
            }
            targetWs.Row(1).Style.Font.Bold = true;
        }

        var rows = sourceWs.RangeUsed().RowsUsed().Skip(1); 
        int targetRowIndex = 2;

        foreach (var row in rows)
        {
            var cell = row.Cell(column);
            
            if (EvaluateCondition(cell, op, targetValueRaw))
            {
                int maxCol = row.LastCellUsed().Address.ColumnNumber;
                for (int i = 1; i <= maxCol; i++)
                {
                    targetWs.Cell(targetRowIndex, i).Value = row.Cell(i).Value;
                }
                targetRowIndex++;
            }
        }

        context.Workbooks.Add(targetAlias, targetWb);
        Console.WriteLine($"     -> Filter Result: {targetRowIndex - 2} rows kept.");
    }

    //Helper
    private bool EvaluateCondition(IXLCell cell, string op, string targetValueStr)
    {
        bool isCellNumeric = cell.DataType == XLDataType.Number;
        double cellNumber = isCellNumeric ? cell.GetValue<double>() : 0;
        
        bool isTargetNumeric = double.TryParse(targetValueStr, out double targetNumber);

        if (isCellNumeric && isTargetNumeric)
        {
            return op.ToLower() switch
            {
                "equals" or "=" or "==" => Math.Abs(cellNumber - targetNumber) < 0.0001,
                "notequals" or "!=" => Math.Abs(cellNumber - targetNumber) > 0.0001,
                "greaterthan" or ">" => cellNumber > targetNumber,
                "lessthan" or "<" => cellNumber < targetNumber,
                "greaterorequal" or ">=" => cellNumber >= targetNumber,
                "lessorequal" or "<=" => cellNumber <= targetNumber,
                _ => false
            };
        }

        string cellText = cell.GetText();
        
        return op.ToLower() switch
        {
            "equals" or "=" or "==" => string.Equals(cellText, targetValueStr, StringComparison.OrdinalIgnoreCase),
            "notequals" or "!=" => !string.Equals(cellText, targetValueStr, StringComparison.OrdinalIgnoreCase),
            "contains" => cellText.Contains(targetValueStr, StringComparison.OrdinalIgnoreCase),
            _ => false
        };
    }
}