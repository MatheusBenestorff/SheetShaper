using System.Text.Json;
using System.Text.Json.Serialization;
using SheetShaper.Core;

Console.WriteLine("========================================");
Console.WriteLine("   SheetShaper CLI - ETL Engine         ");
Console.WriteLine("========================================");

// Folders Config
string rootFolder = Environment.GetEnvironmentVariable("APP_ROOT") ?? "examples";
string configFolder = Path.Combine(rootFolder, "config");
string inputsFolder = Path.Combine(rootFolder, "input");
string outputsFolder = Path.Combine(rootFolder, "output");

// Args
string configPath = args.Length > 0 ? args[0] : Path.Combine(configFolder, "pipeline.json");

if (!File.Exists(configPath))
{
    Console.Error.WriteLine($"[Error] Pipeline config not found at: {configPath}");
    Environment.Exit(1);
}

if (!Directory.Exists(outputsFolder)) Directory.CreateDirectory(outputsFolder);

try
{
    // Load Config
    string jsonContent = File.ReadAllText(configPath);
    
    var jsonOptions = new JsonSerializerOptions
    {
        PropertyNameCaseInsensitive = true,
        WriteIndented = true
    };

    var job = JsonSerializer.Deserialize<SheetJob>(jsonContent, jsonOptions);

    if (job == null || job.Steps.Count == 0)
        throw new Exception("Invalid or empty pipeline configuration.");

    // Execute
    var engine = new SheetEngine();
    engine.Execute(job, inputsFolder, outputsFolder);


    // Results
    Console.WriteLine("----------------------------------------");
    Console.WriteLine($"[SUCCESS] Job '{job.JobName}' completed");
    Console.WriteLine($"Check output folder: {outputsFolder}");
    Console.WriteLine("----------------------------------------");
}
catch (Exception ex)
{
    Console.Error.WriteLine($"[FATAL] Job Failed: {ex.Message}");
    Console.Error.WriteLine(ex.StackTrace); 
}