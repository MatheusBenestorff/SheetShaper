using System;
using System.IO;
using ClosedXML.Excel;
using SheetShaper.Core;

namespace SheetShaper.Tests.Helpers;

public abstract class SheetTestBase : IDisposable
{
    protected readonly string _testRoot;
    protected readonly string _inputFolder;
    protected readonly string _outputFolder;
    protected readonly SheetEngine _engine;

    public SheetTestBase()
    {
        // Setup 
        _testRoot = Path.Combine(Path.GetTempPath(), $"SheetShaper_{Guid.NewGuid()}");
        _inputFolder = Path.Combine(_testRoot, "input");
        _outputFolder = Path.Combine(_testRoot, "output");

        Directory.CreateDirectory(_inputFolder);
        Directory.CreateDirectory(_outputFolder);
        
        _engine = new SheetEngine();
    }

    public void Dispose()
    {
        if (Directory.Exists(_testRoot))
            Directory.Delete(_testRoot, true);
    }

    // --- Path Helpers ---
    protected string GetInPath(string fileName) => Path.Combine(_inputFolder, fileName);
    protected string GetOutPath(string fileName) => Path.Combine(_outputFolder, fileName);

    // --- Seed Helpers ---
    protected void SeedSimpleExcel(string fileName, string content)
    {
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");
        ws.Cell("A1").Value = content;
        wb.SaveAs(GetInPath(fileName));
    }

    protected void SeedMessyExcel(string fileName)
    {
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");
        ws.Cell("A1").Value = "Old_Date"; ws.Cell("B1").Value = "Garbage"; ws.Cell("C1").Value = "Old_Total";
        ws.Cell("A2").Value = "2024-01-01"; ws.Cell("B2").Value = "IgnoreMe"; ws.Cell("C2").Value = "1500";
        wb.SaveAs(GetInPath(fileName));
    }

    protected void SeedFilterDataExcel(string fileName)
    {
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");
        ws.Cell("A1").Value = "Category"; ws.Cell("B1").Value = "Value";
        ws.Cell("A2").Value = "Keep"; ws.Cell("B2").Value = 500;
        ws.Cell("A3").Value = "Drop"; ws.Cell("B3").Value = 100;
        ws.Cell("A4").Value = "Keep"; ws.Cell("B4").Value = 10;
        wb.SaveAs(GetInPath(fileName));
    }

    protected void SeedPivotSourceExcel(string fileName)
    {
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("DadosSAP");
        
        ws.Cell("A1").Value = "Caracteristica"; 
        ws.Cell("B1").Value = "Valor"; 
        ws.Cell("E1").Value = "OV_ID";

        ws.Cell("A2").Value = "Peso"; ws.Cell("B2").Value = "10kg"; ws.Cell("E2").Value = "OV-100";
        ws.Cell("A3").Value = "Cor";  ws.Cell("B3").Value = "Azul"; ws.Cell("E3").Value = "OV-100";

        ws.Cell("A4").Value = "Cor";  ws.Cell("B4").Value = "Verde"; ws.Cell("E4").Value = "OV-200";

        ws.Cell("A5").Value = "Tamanho"; ws.Cell("B5").Value = "G";      ws.Cell("E5").Value = "OV-300";
        ws.Cell("A6").Value = "Peso";    ws.Cell("B6").Value = "5kg";    ws.Cell("E6").Value = "OV-300";
        ws.Cell("A7").Value = "Cor";     ws.Cell("B7").Value = "Branco"; ws.Cell("E7").Value = "OV-300";

        wb.SaveAs(GetInPath(fileName));
    }

    protected void SeedAggregateDataExcel(string fileName)
    {
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("SalesData");
        
        ws.Cell("A1").Value = "Region"; 
        ws.Cell("B1").Value = "SalesPerson"; 
        ws.Cell("C1").Value = "Amount"; 
        ws.Cell("D1").Value = "Score";

        ws.Cell("A2").Value = "North"; ws.Cell("B2").Value = "John"; ws.Cell("C2").Value = 100; ws.Cell("D2").Value = 10;
        ws.Cell("A3").Value = "North"; ws.Cell("B3").Value = "Doe";  ws.Cell("C3").Value = 200; ws.Cell("D3").Value = 20;

        ws.Cell("A4").Value = "South"; ws.Cell("B4").Value = "Jane"; ws.Cell("C4").Value = 500; ws.Cell("D4").Value = 5;
        
        ws.Cell("A5").Value = "East";  ws.Cell("B5").Value = "Bob";  ws.Cell("C5").Value = 0;   ws.Cell("D5").Value = 0;

        wb.SaveAs(GetInPath(fileName));
    }
}