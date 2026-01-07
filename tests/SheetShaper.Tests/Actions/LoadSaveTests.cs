using Xunit;
using System.IO;
using ClosedXML.Excel;
using SheetShaper.Tests.Helpers; 

namespace SheetShaper.Tests.Actions;

public class LoadSaveTests : SheetTestBase
{
    [Fact]
    public void Should_Load_Excel_And_Save_Copy()
    {
        // 1. Arrange
        string inputFile = "source.xlsx";
        string outputFile = "result.xlsx";
        
        SeedSimpleExcel(inputFile, "Hello World");

        var job = new JobBuilder("Test_LoadSave")
            .AddLoadStep("1", inputFile, "MySheet")
            .AddSaveStep("2", outputFile, "MySheet")
            .Build();

        // 2. Act
        _engine.Execute(job, _inputFolder, _outputFolder);

        // 3. Assert
        Assert.True(File.Exists(GetOutPath(outputFile)));
        
        using var wb = new XLWorkbook(GetOutPath(outputFile));
        Assert.Equal("Hello World", wb.Worksheet(1).Cell("A1").Value.ToString());
    }

    [Fact]
    public void Should_Throw_Error_When_Input_File_Missing()
    {
        var job = new JobBuilder("Fail_Job")
            .AddLoadStep("1", "ghost.xlsx", "Ghost")
            .Build();

        Assert.Throws<FileNotFoundException>(() => 
            _engine.Execute(job, _inputFolder, _outputFolder)
        );
    }
}