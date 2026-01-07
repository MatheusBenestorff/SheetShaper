using Xunit;
using System.IO;
using ClosedXML.Excel;
using SheetShaper.Tests.Helpers;
using System.Linq;

namespace SheetShaper.Tests.Actions;

public class FilterRowsTests : SheetTestBase
{
    [Fact]
    public void Should_Filter_Rows_Correctly()
    {
        string inputFile = "filter.xlsx";
        string outputFile = "filtered.xlsx";
        
        SeedFilterDataExcel(inputFile);

        var job = new JobBuilder("Test_Filtering")
            .AddLoadStep("1", inputFile, "Source")
            .AddFilterStep("2", "Source", "Step1", "A", "Equals", "Keep")
            .AddFilterStep("3", "Step1", "Final", "B", "GreaterThan", 100)
            .AddSaveStep("4", outputFile, "Final")
            .Build();

        _engine.Execute(job, _inputFolder, _outputFolder);

        using var wb = new XLWorkbook(GetOutPath(outputFile));
        var ws = wb.Worksheet(1);
        
        int rowCount = ws.RangeUsed().RowsUsed().Count();
        Assert.Equal(2, rowCount); 
        Assert.Equal(500, ws.Cell("B2").GetValue<double>());
    }
}