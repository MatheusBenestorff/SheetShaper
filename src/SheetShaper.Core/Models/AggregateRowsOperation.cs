namespace SheetShaper.Core;

// Auxiliary class for JSON AggregateRows operations
public class AggregateRowsOperation
{
    public string TargetColumn { get; set; }
    public string SourceColumn { get; set; }
    public string Type { get; set; }
}