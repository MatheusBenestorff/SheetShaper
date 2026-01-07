using Xunit;
using System.Collections.Generic;
using System.IO;
using ClosedXML.Excel;
using SheetShaper.Tests.Helpers;

namespace SheetShaper.Tests.Actions;

public class MapColumnsTests : SheetTestBase
{
    [Fact]
    public void Should_Map_Columns_Correctly()
    {
        string inputFile = "messy.xlsx";
        string outputFile = "clean.xlsx";

        SeedMessyExcel(inputFile);

        var mappingRules = new List<object>
        {
            new { From = "A", To = "A", Header = "Date" },    
            new { From = "C", To = "B", Header = "Revenue" }    
        };

        var job = new JobBuilder("Test_Mapping")
            .AddLoadStep("1", inputFile, "Source")
            .AddMapStep("2", "Source", "Report", mappingRules)
            .AddSaveStep("3", outputFile, "Report")
            .Build();

        _engine.Execute(job, _inputFolder, _outputFolder);

        Assert.True(File.Exists(GetOutPath(outputFile)));
        
        using var wb = new XLWorkbook(GetOutPath(outputFile));
        var ws = wb.Worksheet(1);
        Assert.Equal("Date", ws.Cell("A1").Value.ToString());
        Assert.Equal("Revenue", ws.Cell("B1").Value.ToString());
        Assert.True(ws.Cell("C1").IsEmpty());
    }
}