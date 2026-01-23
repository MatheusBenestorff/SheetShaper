using Xunit;
using System.IO;
using ClosedXML.Excel;
using SheetShaper.Tests.Helpers;
using System.Collections.Generic;

namespace SheetShaper.Tests.Actions;

public class AggregateRowsTests : SheetTestBase
{
    [Fact]
    public void Should_Aggregate_And_Calculate_Values_Correctly()
    {
        // ARRANGE
        string inputFile = "sales_transactions.xlsx";
        string outputFile = "sales_summary.xlsx";
        
        SeedAggregateDataExcel(inputFile);

        var operations = new List<object>
        {
            new { targetColumn = "Total Amount", sourceColumn = "C", type = "Sum" },   
            new { targetColumn = "Avg Score",    sourceColumn = "D", type = "Average" },
            new { targetColumn = "Tx Count",     type = "Count" }                       
        };

        var job = new JobBuilder("Test_Aggregation")
            .AddLoadStep("1", inputFile, "Source")
            .AddAggregateStep("2", "Source", "Summary", 
                              groupByCol: "A", 
                              operations: operations)
            .AddSaveStep("3", outputFile, "Summary")
            .Build();

        // ACT
        _engine.Execute(job, _inputFolder, _outputFolder);

        // ASSERT
        Assert.True(File.Exists(GetOutPath(outputFile)));

        using var wb = new XLWorkbook(GetOutPath(outputFile));
        var ws = wb.Worksheet(1);

        Assert.Equal("Group Key",    ws.Cell("A1").Value.ToString());
        Assert.Equal("Total Amount", ws.Cell("B1").Value.ToString());
        Assert.Equal("Avg Score",    ws.Cell("C1").Value.ToString());
        Assert.Equal("Tx Count",     ws.Cell("D1").Value.ToString());

        Assert.Equal("North", ws.Cell("A2").Value.ToString());
        
        Assert.Equal(300, ws.Cell("B2").GetValue<double>());
        
        Assert.Equal(15, ws.Cell("C2").GetValue<double>());
        
        Assert.Equal(2, ws.Cell("D2").GetValue<double>());

        Assert.Equal("South", ws.Cell("A3").Value.ToString());
        Assert.Equal(500, ws.Cell("B3").GetValue<double>()); 
        Assert.Equal(5,   ws.Cell("C3").GetValue<double>()); 
        Assert.Equal(1,   ws.Cell("D3").GetValue<double>()); 
    }
}