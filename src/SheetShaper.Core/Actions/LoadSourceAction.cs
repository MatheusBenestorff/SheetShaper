using ClosedXML.Excel;

namespace SheetShaper.Core.Actions;

public class LoadSourceAction : ISheetAction
{
    public string ActionName => "LoadSource";

    public void Execute(ActionContext context)
    {
        string fileName = context.GetParam("file");
        string alias = context.GetParam("alias");
        
        string fullPath = Path.Combine(context.InputFolder, fileName);

        if (!File.Exists(fullPath))
            throw new FileNotFoundException($"Input Excel file not found: {fullPath}");

        Console.WriteLine($"     loading '{fileName}' as '{alias}'...");
        
        var workbook = new XLWorkbook(fullPath);
        context.Workbooks.Add(alias, workbook);    }
}
