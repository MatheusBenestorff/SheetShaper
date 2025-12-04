namespace SheetShaper.Core;

public class SheetJob
{
    public string JobName { get; set; } = string.Empty;
    public List<PipelineStep> Steps { get; set; } = new();
}