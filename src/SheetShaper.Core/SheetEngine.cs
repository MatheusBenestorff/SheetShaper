using ClosedXML.Excel;
using SheetShaper.Core.Actions; 

namespace SheetShaper.Core;

public class SheetEngine
{
    private readonly Dictionary<string, IXLWorkbook> _workbooks = new();
    
    // Register of all actions
    private readonly Dictionary<string, ISheetAction> _availableActions;
    
    public SheetEngine()
    {
        _availableActions = new Dictionary<string, ISheetAction>(StringComparer.OrdinalIgnoreCase);

        Register(new LoadSourceAction());
        Register(new SaveFileAction());
        Register(new MapColumnsAction());
        Register(new FilterRowsAction()); 
    }

    private void Register(ISheetAction action)
    {
        _availableActions[action.ActionName] = action;
    }

    public void Execute(SheetJob job, string inputRoot, string outputRoot)
    {
        Console.WriteLine($"[Engine] Starting Job: {job.JobName}");

        try
        {
            foreach (var step in job.Steps)
            {
                Console.WriteLine($"  -> Executing Step {step.StepId}: {step.Action}");

                if (!_availableActions.ContainsKey(step.Action))
                {
                    Console.WriteLine($"     [Warn] Unknown Action: {step.Action}");
                    continue;
                }

                var action = _availableActions[step.Action];
                var context = new ActionContext(_workbooks, step, inputRoot, outputRoot);

                try
                {
                    action.Execute(context);
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine($"[Error] Step {step.StepId} ({step.Action}) failed: {ex.Message}");
                    throw;
                }
            }
        }
        finally
        {
            foreach(var wb in _workbooks.Values) wb.Dispose();
            _workbooks.Clear();
        }
    }
}