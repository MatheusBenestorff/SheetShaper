using Xunit;
using SheetShaper.Core;
using ClosedXML.Excel;
using System.IO;
using System.Collections.Generic;
using System;

namespace SheetShaper.Tests;

public class SheetEngineTests : IDisposable
{
    // Temporary folders
    private readonly string _testRoot;
    private readonly string _inputFolder;
    private readonly string _outputFolder;

    public SheetEngineTests()
    {
        _testRoot = Path.Combine(Path.GetTempPath(), $"SheetShaper_{Guid.NewGuid()}");
        _inputFolder = Path.Combine(_testRoot, "input");
        _outputFolder = Path.Combine(_testRoot, "output");

        Directory.CreateDirectory(_inputFolder);
        Directory.CreateDirectory(_outputFolder);
    }

    // Teardown: Clean up
    public void Dispose()
    {
        if (Directory.Exists(_testRoot))
        {
            Directory.Delete(_testRoot, true);
        }
    }

    [Fact]
    public void Should_Load_Excel_And_Save_Copy()
    {
        // ARRANGE
        string inputFileName = "source.xlsx";
        string outputFileName = "result.xlsx";
        
        CreateDummyExcel(Path.Combine(_inputFolder, inputFileName), "Hello World");

        var job = new SheetJob
        {
            JobName = "Test_Integration",
            Steps = new List<PipelineStep>
            {
                new PipelineStep
                {
                    StepId = "1",
                    Action = "LoadSource",
                    Params = new Dictionary<string, object>
                    {
                        { "file", inputFileName },
                        { "alias", "MySheet" }
                    }
                },
                new PipelineStep
                {
                    StepId = "2",
                    Action = "SaveFile",
                    Params = new Dictionary<string, object>
                    {
                        { "fileName", outputFileName },
                        { "sourceAlias", "MySheet" }
                    }
                }
            }
        };

        var engine = new SheetEngine();

        // ACT 
        engine.Execute(job, _inputFolder, _outputFolder);

        // ASSERT 
        string expectedOutputPath = Path.Combine(_outputFolder, outputFileName);
        
        Assert.True(File.Exists(expectedOutputPath), "Output file should exist.");

        using var wb = new XLWorkbook(expectedOutputPath);
        var cellValue = wb.Worksheet(1).Cell("A1").Value.ToString();
        Assert.Equal("Hello World", cellValue);
    }

    [Fact]
    public void Should_Throw_Error_When_Input_File_Missing()
    {
        // ARRANGE
        var job = new SheetJob
        {
            JobName = "Fail_Job",
            Steps = new List<PipelineStep>
            {
                new PipelineStep
                {
                    StepId = "1",
                    Action = "LoadSource",
                    Params = new Dictionary<string, object>
                    {
                        { "file", "ghost_file.xlsx" }, // File dont exist
                        { "alias", "Ghost" }
                    }
                }
            }
        };

        var engine = new SheetEngine();

        // ACT & ASSERT
        Assert.Throws<FileNotFoundException>(() => 
            engine.Execute(job, _inputFolder, _outputFolder)
        );
    }

    [Fact]
    public void Should_Map_Columns_Correctly()
    {
        //ARRANGE
        string inputFile = "messy_data.xlsx";
        string outputFile = "clean_data.xlsx";
        string inputPath = Path.Combine(_inputFolder, inputFile);

        using (var wb = new XLWorkbook())
        {
            var ws = wb.Worksheets.Add("Sheet1");
            
            ws.Cell("A1").Value = "Old_Date";
            ws.Cell("A2").Value = "2024-01-01";

            ws.Cell("B1").Value = "Garbage_Data";
            ws.Cell("B2").Value = "IgnoreMe";

            ws.Cell("C1").Value = "Old_Total";
            ws.Cell("C2").Value = "1500";

            wb.SaveAs(inputPath);
        }

        var mappingRules = new List<object>
        {
            new { From = "A", To = "A", Header = "Date" },    
            new { From = "C", To = "B", Header = "Revenue" }    
        };

        string mappingsJson = System.Text.Json.JsonSerializer.Serialize(mappingRules);

        var job = new SheetJob
        {
            JobName = "Test_Mapping",
            Steps = new List<PipelineStep>
            {
                new PipelineStep
                {
                    StepId = "1", Action = "LoadSource",
                    Params = new Dictionary<string, object> 
                    { 
                        { "file", inputFile }, 
                        { "alias", "Source" } 
                    }
                },
                new PipelineStep
                {
                    StepId = "2", Action = "MapColumns",
                    Params = new Dictionary<string, object> 
                    { 
                        { "sourceSheet", "Source" },
                        { "targetSheet", "CleanReport" },
                        { "mappings", mappingsJson } 
                    }
                },
                new PipelineStep
                {
                    StepId = "3", Action = "SaveFile",
                    Params = new Dictionary<string, object> 
                    { 
                        { "fileName", outputFile }, 
                        { "sourceAlias", "CleanReport" } 
                    }
                }
            }
        };

        var engine = new SheetEngine();

        //ACT
        engine.Execute(job, _inputFolder, _outputFolder);

        //ASSERT
        string outputPath = Path.Combine(_outputFolder, outputFile);
        Assert.True(File.Exists(outputPath), "Output file should have been created.");

        using var resultWb = new XLWorkbook(outputPath);
        var wsResult = resultWb.Worksheet(1);

        Assert.Equal("Date", wsResult.Cell("A1").Value.ToString());
        Assert.Equal("2024-01-01", wsResult.Cell("A2").Value.ToString());

        Assert.Equal("Revenue", wsResult.Cell("B1").Value.ToString());
        Assert.Equal("1500", wsResult.Cell("B2").Value.ToString());

        Assert.True(wsResult.Cell("C1").IsEmpty());
    }

    [Fact]
    public void Should_Filter_Rows_Correctly()
    {
        // ARRANGE
        string inputFile = "filter_test.xlsx";
        string outputFile = "filtered_output.xlsx";
        
        using (var wb = new XLWorkbook())
        {
            var ws = wb.Worksheets.Add("Sheet1");
            ws.Cell("A1").Value = "Category"; ws.Cell("B1").Value = "Value";
            
            ws.Cell("A2").Value = "Keep";     ws.Cell("B2").Value = 500; 
            ws.Cell("A3").Value = "Drop";     ws.Cell("B3").Value = 100;
            ws.Cell("A4").Value = "Keep";     ws.Cell("B4").Value = 10; 
            
            wb.SaveAs(Path.Combine(_inputFolder, inputFile));
        }

        var job = new SheetJob
        {
            JobName = "Test_Filtering",
            Steps = new List<PipelineStep>
            {
                new PipelineStep { StepId="1", Action="LoadSource", Params=new Dictionary<string,object> { {"file", inputFile}, {"alias", "Source"} } },
                
                new PipelineStep 
                { 
                    StepId="2", Action="FilterRows", 
                    Params=new Dictionary<string,object> 
                    { 
                        {"sourceSheet", "Source"}, {"targetSheet", "Step1"}, 
                        {"column", "A"}, {"operator", "Equals"}, {"value", "Keep"} 
                    } 
                },

                new PipelineStep 
                { 
                    StepId="3", Action="FilterRows", 
                    Params=new Dictionary<string,object> 
                    { 
                        {"sourceSheet", "Step1"}, {"targetSheet", "Final"}, 
                        {"column", "B"}, {"operator", "GreaterThan"}, {"value", 100} 
                    } 
                },

                new PipelineStep { StepId="4", Action="SaveFile", Params=new Dictionary<string,object> { {"fileName", outputFile}, {"sourceAlias", "Final"} } }
            }
        };

        var engine = new SheetEngine();

        //ACT
        engine.Execute(job, _inputFolder, _outputFolder);

        // ASSERT
        using var resultWb = new XLWorkbook(Path.Combine(_outputFolder, outputFile));
        var wsResult = resultWb.Worksheet(1);
        
        int rowCount = wsResult.RangeUsed().RowsUsed().Count();
        Assert.Equal(2, rowCount); 

        Assert.Equal(500, wsResult.Cell("B2").GetValue<double>());
    }

    // Aux
    private void CreateDummyExcel(string path, string content)
    {
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");
        ws.Cell("A1").Value = content;
        wb.SaveAs(path);
    }
}