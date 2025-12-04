using ClosedXML.Excel;
using System.Text.Json;

namespace SheetShaper.Core;

public class SheetEngine
{
    private readonly Dictionary<string, IXLWorkbook> _workbookContext = new();

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
        
        foreach(var wb in _workbookContext.Values) wb.Dispose();
        _workbookContext.Clear();
    }

    private void ExecuteStep(PipelineStep step, string inPath, string outPath)
    {
        switch (step.Action)
        {
            case "LoadSource":
                // Implement Load
                break;
            case "MapColumns":
                // Implement Mapping
                break;
            case "SaveFile":
                // Implement Save
                break;
            default:
                Console.WriteLine($"     [Warn] Unknown Action: {step.Action}");
                break;
        }
    }
}