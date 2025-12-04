namespace SheetShaper.Core;

public class PipelineStep
{
    public string StepId { get; set; } = string.Empty;
    
    public string Action { get; set; } = string.Empty; 
    
    public Dictionary<string, object> Params { get; set; } = new();
}