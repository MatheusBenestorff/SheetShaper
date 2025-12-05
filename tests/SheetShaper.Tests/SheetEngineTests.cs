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

    // Aux
    private void CreateDummyExcel(string path, string content)
    {
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");
        ws.Cell("A1").Value = content;
        wb.SaveAs(path);
    }
}