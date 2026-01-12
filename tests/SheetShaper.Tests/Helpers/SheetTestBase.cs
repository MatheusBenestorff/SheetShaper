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
        
        // Cabeçalhos (A=Caracteristica, B=Valor, E=NumeroOV)
        ws.Cell("A1").Value = "Caracteristica"; 
        ws.Cell("B1").Value = "Valor"; 
        ws.Cell("E1").Value = "OV_ID";

        // OV-100 (Tem Peso e Cor)
        ws.Cell("A2").Value = "Peso"; ws.Cell("B2").Value = "10kg"; ws.Cell("E2").Value = "OV-100";
        ws.Cell("A3").Value = "Cor";  ws.Cell("B3").Value = "Azul"; ws.Cell("E3").Value = "OV-100";

        // OV-200 (Tem apenas Cor)
        ws.Cell("A4").Value = "Cor";  ws.Cell("B4").Value = "Verde"; ws.Cell("E4").Value = "OV-200";

        // OV-300 (Tem Peso, Cor e Tamanho)
        ws.Cell("A5").Value = "Tamanho"; ws.Cell("B5").Value = "G";      ws.Cell("E5").Value = "OV-300";
        ws.Cell("A6").Value = "Peso";    ws.Cell("B6").Value = "5kg";    ws.Cell("E6").Value = "OV-300";
        ws.Cell("A7").Value = "Cor";     ws.Cell("B7").Value = "Branco"; ws.Cell("E7").Value = "OV-300";

        wb.SaveAs(GetInPath(fileName));
    }
}