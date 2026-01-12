using ClosedXML.Excel;

namespace SheetShaper.Core.Actions;

public class PivotSheetAction : ISheetAction
{
    public string ActionName => "PivotSheet";

    public void Execute(ActionContext context)
    {
        // Parameters
        string sourceAlias = context.GetParam("sourceSheet");
        string targetAlias = context.GetParam("targetSheet");
        string groupByCol = context.GetParam("groupByColumn");   
        string keyCol = context.GetParam("pivotKeyColumn");      
        string valCol = context.GetParam("pivotValueColumn");    

        if (!context.Workbooks.ContainsKey(sourceAlias))
            throw new Exception($"Source workbook '{sourceAlias}' not loaded.");

        var sourceWs = context.Workbooks[sourceAlias].Worksheets.First();
        var targetWb = new XLWorkbook();
        var targetWs = targetWb.Worksheets.Add("PivotedData");

        Console.WriteLine($"     Pivoting '{sourceAlias}': GroupBy {groupByCol}, Keys from {keyCol}...");

        // Data Structure in Memory
        var groupedData = new Dictionary<string, Dictionary<string, string>>();
        
        var allPossibleKeys = new HashSet<string>();

        var rows = sourceWs.RangeUsed().RowsUsed().Skip(1); 

        // Reading Phase
        foreach (var row in rows)
        {
            string groupID = row.Cell(groupByCol).GetString().Trim();
            string keyName = row.Cell(keyCol).GetString().Trim();     
            string value = row.Cell(valCol).GetString().Trim();       

            if (string.IsNullOrEmpty(groupID)) continue;

            if (!groupedData.ContainsKey(groupID))
            {
                groupedData[groupID] = new Dictionary<string, string>();
            }

            groupedData[groupID][keyName] = value;
            
            allPossibleKeys.Add(keyName);
        }

        var sortedColumns = allPossibleKeys.OrderBy(x => x).ToList();

        // Writing Phase
        
        targetWs.Cell(1, 1).Value = "ID"; // First Column
        targetWs.Cell(1, 1).Style.Font.Bold = true;

        for (int i = 0; i < sortedColumns.Count; i++)
        {
            targetWs.Cell(1, i + 2).Value = sortedColumns[i]; 
            targetWs.Cell(1, i + 2).Style.Font.Bold = true;
        }

        int currentRow = 2;
        foreach (var group in groupedData)
        {
            targetWs.Cell(currentRow, 1).Value = group.Key;

            for (int i = 0; i < sortedColumns.Count; i++)
            {
                string colName = sortedColumns[i];
                
                if (group.Value.ContainsKey(colName))
                {
                    targetWs.Cell(currentRow, i + 2).Value = group.Value[colName];
                }
            }
            currentRow++;
        }

        context.Workbooks.Add(targetAlias, targetWb);
        Console.WriteLine($"     -> Pivot Result: {groupedData.Count} unique IDs found with {sortedColumns.Count} distinct characteristics.");
    }
}