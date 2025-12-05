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
                // Implement Mapping
                break;
            case "SaveFile":
                // ExecuteSaveFile(step, outPath);
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



    // --- HELPERS ---

    private string GetParam(PipelineStep step, string key)
    {
        if (!step.Params.ContainsKey(key))
            throw new ArgumentException($"Missing required parameter '{key}' for action '{step.Action}'");

        var element = (JsonElement)step.Params[key];
        return element.ToString();
    }

    private string? GetOptionalParam(PipelineStep step, string key)
    {
        if (!step.Params.ContainsKey(key)) return null;
        var element = (JsonElement)step.Params[key];
        return element.ToString();
    }
}