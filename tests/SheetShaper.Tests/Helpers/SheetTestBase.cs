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
}