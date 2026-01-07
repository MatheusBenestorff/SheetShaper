namespace SheetShaper.Core.Actions;

public interface ISheetAction
{
    // The name that goes on the JSON pipeline
    string ActionName { get; }
    
    void Execute(ActionContext context);
}