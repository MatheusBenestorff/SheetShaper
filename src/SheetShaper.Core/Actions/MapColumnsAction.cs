using ClosedXML.Excel;
using System.Text.Json;

namespace SheetShaper.Core.Actions;

public class MapColumnsAction : ISheetAction
{
    public string ActionName => "MapColumns";

    public void Execute(ActionContext context)
    {
        string sourceAlias = context.GetParam("sourceSheet");
        string targetAlias = context.GetParam("targetSheet");

        if (! context.Workbooks.ContainsKey(sourceAlias))
            throw new Exception($"Source workbook '{sourceAlias}' not loaded.");

        var sourceWb = context.Workbooks[sourceAlias];
        var sourceWs = sourceWb.Worksheets.First();

        var targetWb = new XLWorkbook();
        var targetWs = targetWb.Worksheets.Add("Data");
        
        Console.WriteLine($"     Mapping columns from '{sourceAlias}' to new workbook '{targetAlias}'...");

        var mappingsJson = context.GetParam("mappings");
        var rules = JsonSerializer.Deserialize<List<MappingRule>>(mappingsJson, new JsonSerializerOptions { PropertyNameCaseInsensitive = true });

        if (rules == null || !rules.Any())
            throw new Exception("No mappings defined.");

        int lastRow = sourceWs.LastRowUsed()?.RowNumber() ?? 0;

        foreach (var rule in rules)
        {
            targetWs.Cell($"{rule.To}1").Value = rule.Header;
            targetWs.Cell($"{rule.To}1").Style.Font.Bold = true; 

            if (lastRow >= 2)
            {
                var sourceRange = sourceWs.Range($"{rule.From}2:{rule.From}{lastRow}");
                
                var targetCell = targetWs.Cell($"{rule.To}2");
                
                var dataValues = sourceRange.Cells().Select(c => c.Value).ToList();
                
                for (int i = 0; i < dataValues.Count; i++)
                {
                    targetWs.Cell(2 + i, rule.To).Value = dataValues[i];
                }
            }
        }

        context.Workbooks.Add(targetAlias, targetWb);
    }
}