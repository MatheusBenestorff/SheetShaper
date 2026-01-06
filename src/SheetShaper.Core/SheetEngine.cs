using ClosedXML.Excel;
using System.Text.Json;

namespace SheetShaper.Core;

public class SheetEngine
{
    private readonly Dictionary<string, IXLWorkbook> _workbooks = new();

    public void Execute(SheetJob job, string inputRootFolder, string outputRootFolder)
    {
        Console.WriteLine($"[Engine] Starting Job: {job.JobName}");

        foreach (var step in job.Steps)
        {
            Console.WriteLine($"  -> Executing Step {step.StepId}: {step.Action}");
            
            try 
            {
                ExecuteStep(step, inputRootFolder, outputRootFolder);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"[Error] Step {step.StepId} failed: {ex.Message}");
                throw; 
            }
        }
        
        foreach(var wb in _workbooks.Values) 
        {
            wb.Dispose();
        }
        _workbooks.Clear();
    }

    private void ExecuteStep(PipelineStep step, string inPath, string outPath)
    {
        switch (step.Action)
        {
            case "LoadSource":
                ExecuteLoadSource(step, inPath);
                break;
            case "MapColumns":
                ExecuteMapColumns(step);
                break;
            case "SaveFile":
                ExecuteSaveFile(step, outPath);
                break;
            default:
                Console.WriteLine($"     [Warn] Unknown Action: {step.Action}");
                break;
        }
    }

    // --- ACTIONS ---

    private void ExecuteLoadSource(PipelineStep step, string inputRoot)
    {
        string fileName = GetParam(step, "file");
        string alias = GetParam(step, "alias");
        
        string fullPath = Path.Combine(inputRoot, fileName);

        if (!File.Exists(fullPath))
            throw new FileNotFoundException($"Input Excel file not found: {fullPath}");

        Console.WriteLine($"     loading '{fileName}' as '{alias}'...");
        
        var workbook = new XLWorkbook(fullPath);
        _workbooks.Add(alias, workbook);
    }

    private void ExecuteMapColumns(PipelineStep step)
    {
        string sourceAlias = GetParam(step, "sourceSheet");
        string targetAlias = GetParam(step, "targetSheet");
        
        if (!_workbooks.ContainsKey(sourceAlias))
            throw new Exception($"Source workbook '{sourceAlias}' not loaded.");

        var sourceWb = _workbooks[sourceAlias];
        var sourceWs = sourceWb.Worksheets.First();

        var targetWb = new XLWorkbook();
        var targetWs = targetWb.Worksheets.Add("Data");
        
        Console.WriteLine($"     Mapping columns from '{sourceAlias}' to new workbook '{targetAlias}'...");

        var mappingsJson = GetParam(step, "mappings");
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

        _workbooks.Add(targetAlias, targetWb);
    }

    private void ExecuteSaveFile(PipelineStep step, string outputRoot)
    {
        string fileName = GetParam(step, "fileName");
        
        string alias = GetOptionalParam(step, "sourceAlias");

        IXLWorkbook workbookToSave;

        if (string.IsNullOrEmpty(alias))
        {
            workbookToSave = _workbooks.Values.First();
        }
        else
        {
            if (!_workbooks.ContainsKey(alias))
                throw new Exception($"Workbook with alias '{alias}' not found in memory.");
            workbookToSave = _workbooks[alias];
        }

        string fullPath = Path.Combine(outputRoot, fileName);
        Console.WriteLine($"     saving to '{fileName}'...");

        workbookToSave.SaveAs(fullPath);
    }


    // --- HELPERS ---

    private string GetParam(PipelineStep step, string key)
    {
        if (!step.Params.ContainsKey(key))
            throw new ArgumentException($"Missing required parameter '{key}' for action '{step.Action}'");

        object value = step.Params[key];

        if (value is JsonElement element)
        {
            return element.ToString();
        }

        return value.ToString() ?? string.Empty;
    }

    private string? GetOptionalParam(PipelineStep step, string key)
    {
        if (!step.Params.ContainsKey(key)) return null;
        
        object value = step.Params[key];

        if (value is JsonElement element)
        {
            return element.ToString();
        }

        return value.ToString();
    }
}