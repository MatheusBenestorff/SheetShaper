using ClosedXML.Excel;
using SheetShaper.Core; 

namespace SheetShaper.Core.Actions;

public class SaveFileAction : ISheetAction
{
    public string ActionName => "SaveFile";

    public void Execute(ActionContext context)
    {
        string fileName = context.GetParam("fileName");
        
        string? alias = context.GetOptionalParam("sourceAlias");

        IXLWorkbook workbookToSave;

        if (string.IsNullOrEmpty(alias))
        {
            if (!context.Workbooks.Any())
                throw new Exception("No workbooks available to save.");
                
            workbookToSave = context.Workbooks.Values.First();
        }
        else
        {
            if (!context.Workbooks.ContainsKey(alias))
                throw new Exception($"Workbook with alias '{alias}' not found in memory.");
                
            workbookToSave = context.Workbooks[alias];
        }

        string fullPath = Path.Combine(context.OutputFolder, fileName);
        
        Console.WriteLine($"     saving to '{fileName}'...");

        workbookToSave.SaveAs(fullPath);
    }
}