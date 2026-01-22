using ClosedXML.Excel;
using System.Text.Json;

namespace SheetShaper.Core.Actions;

public class AggregateRowsAction : ISheetAction
{
    public string ActionName => "AggregateRows";

    public void Execute(ActionContext context)
    {
        //Inputs
        string sourceAlias = context.GetParam("sourceSheet");
        string targetAlias = context.GetParam("targetSheet");
        string groupByCol = context.GetParam("groupByColumn");
        
        // Operations come as JSON Array
        string opsJson = context.GetParam("operations");
        var operations = JsonSerializer.Deserialize<List<AggregateRowsOperation>>(opsJson, new JsonSerializerOptions { PropertyNameCaseInsensitive = true });

        if (!context.Workbooks.ContainsKey(sourceAlias))
            throw new Exception($"Source workbook '{sourceAlias}' not loaded.");

        var sourceWs = context.Workbooks[sourceAlias].Worksheets.First();
        var targetWb = new XLWorkbook();
        var targetWs = targetWb.Worksheets.Add("Summary");

        Console.WriteLine($"     Aggregating '{sourceAlias}' by Column {groupByCol}...");

        var rows = sourceWs.RangeUsed().RowsUsed().Skip(1); 
        
        // Group the rows using a Dictionary where Key = groupByColumn Value
        var groups = rows.GroupBy(r => r.Cell(groupByCol).GetString().Trim());

        targetWs.Cell(1, 1).Value = "Group Key"; 
        for (int i = 0; i < operations.Count; i++)
        {
            targetWs.Cell(1, i + 2).Value = operations[i].TargetColumn;
        }
        targetWs.Row(1).Style.Font.Bold = true;

        // Processes Calculations
        int currentRow = 2;
        foreach (var group in groups)
        {
            // Write the key
            targetWs.Cell(currentRow, 1).Value = group.Key;

            // For each requested operation
            for (int i = 0; i < operations.Count; i++)
            {
                var op = operations[i];
                double result = 0;

                switch (op.Type.ToLower())
                {
                    case "count":
                        result = group.Count();
                        break;

                    case "sum":
                        result = group.Sum(r => GetDouble(r, op.SourceColumn));
                        break;

                    case "average":
                        result = group.Average(r => GetDouble(r, op.SourceColumn));
                        break;

                    case "max":
                        result = group.Max(r => GetDouble(r, op.SourceColumn));
                        break;
                    
                    case "min":
                        result = group.Min(r => GetDouble(r, op.SourceColumn));
                        break;
                }

                targetWs.Cell(currentRow, i + 2).Value = result;
            }
            currentRow++;
        }

        context.Workbooks.Add(targetAlias, targetWb);
        Console.WriteLine($"     -> Aggregation Result: {groups.Count()} unique groups created.");
    }

    // Helper
    private double GetDouble(IXLRangeRow row, string colLetter)
    {
        if (string.IsNullOrEmpty(colLetter)) return 0;
        
        var cell = row.Cell(colLetter);
        if (cell.DataType == XLDataType.Number)
            return cell.GetValue<double>();
            
        if (double.TryParse(cell.GetString(), out double val))
            return val;

        return 0;
    }

}