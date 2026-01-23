using Xunit;
using System.IO;
using ClosedXML.Excel;
using SheetShaper.Tests.Helpers;

namespace SheetShaper.Tests.Actions;

public class AddFormulaColumnTests : SheetTestBase
{
    [Fact]
    public void Should_Add_Formula_And_Format_Correctly()
    {
        // ARRANGE
        string inputFile = "formula_input.xlsx";
        string outputFile = "formula_result.xlsx";
        
        SeedFormulaDataExcel(inputFile);

        var job = new JobBuilder("Test_Formula_Injection")
            .AddLoadStep("1", inputFile, "Source")
            .AddFormulaStep("2", "Source", "Total Value", 
                            formula: "A{row}*B{row}", 
                            format: "$ #,##0.00")    
            .AddSaveStep("3", outputFile, "Source")
            .Build();

        // ACT
        _engine.Execute(job, _inputFolder, _outputFolder);

        // ASSERT
        Assert.True(File.Exists(GetOutPath(outputFile)));

        using var wb = new XLWorkbook(GetOutPath(outputFile));
        var ws = wb.Worksheet(1);

        // Structural Validation

        Assert.Equal("Total Value", ws.Cell("C1").Value.ToString());

        var cellC2 = ws.Cell("C2");
        
        Assert.Equal("A2*B2", cellC2.FormulaA1);
        
        Assert.Equal(55.0, cellC2.GetValue<double>());
        
        Assert.Equal("$ #,##0.00", cellC2.Style.NumberFormat.Format);

        var cellC3 = ws.Cell("C3");
        Assert.Equal("A3*B3", cellC3.FormulaA1);
        Assert.Equal(40.0, cellC3.GetValue<double>());
    }
}