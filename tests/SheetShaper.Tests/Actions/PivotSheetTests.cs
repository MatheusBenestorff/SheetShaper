using Xunit;
using System.IO;
using ClosedXML.Excel;
using SheetShaper.Tests.Helpers;
using System.Linq;

namespace SheetShaper.Tests.Actions;

public class PivotSheetTests : SheetTestBase
{
    [Fact]
    public void Should_Pivot_Rows_Into_Columns_Correctly()
    {
        // 1. ARRANGE
        string inputFile = "sap_data.xlsx";
        string outputFile = "report_ov.xlsx";
        
        SeedPivotSourceExcel(inputFile);

        var job = new JobBuilder("Test_Pivot_SAP")
            .AddLoadStep("1", inputFile, "Source")
            .AddPivotStep("2", "Source", "Pivoted", 
                          groupCol: "E",   
                          keyCol: "A",     
                          valCol: "B")     
            .AddSaveStep("3", outputFile, "Pivoted")
            .Build();

        // 2. ACT
        _engine.Execute(job, _inputFolder, _outputFolder);

        // 3. ASSERT
        Assert.True(File.Exists(GetOutPath(outputFile)));

        using var wb = new XLWorkbook(GetOutPath(outputFile));
        var ws = wb.Worksheet(1);

        
        Assert.Equal("ID (OV)", ws.Cell("A1").Value.ToString());
        Assert.Equal("Cor",     ws.Cell("B1").Value.ToString());
        Assert.Equal("Peso",    ws.Cell("C1").Value.ToString());
        Assert.Equal("Tamanho", ws.Cell("D1").Value.ToString());


        Assert.Equal("OV-100", ws.Cell("A2").Value.ToString());
        Assert.Equal("Azul",   ws.Cell("B2").Value.ToString()); 
        Assert.Equal("10kg",   ws.Cell("C2").Value.ToString());
        Assert.True(ws.Cell("D2").IsEmpty());                   

        Assert.Equal("OV-200", ws.Cell("A3").Value.ToString());
        Assert.Equal("Verde",  ws.Cell("B3").Value.ToString()); 
        Assert.True(ws.Cell("C3").IsEmpty());                  

        Assert.Equal("OV-300", ws.Cell("A4").Value.ToString());
        Assert.Equal("Branco", ws.Cell("B4").Value.ToString());
        Assert.Equal("5kg",    ws.Cell("C4").Value.ToString());
        Assert.Equal("G",      ws.Cell("D4").Value.ToString());
    }
}