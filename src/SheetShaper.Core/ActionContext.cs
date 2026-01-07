using ClosedXML.Excel;
using System.Text.Json;

namespace SheetShaper.Core;

public class ActionContext
{
    public Dictionary<string, IXLWorkbook> Workbooks { get; }
    public PipelineStep Step { get; }
    public string InputFolder { get; }
    public string OutputFolder { get; }

    public ActionContext(Dictionary<string, IXLWorkbook> workbooks, PipelineStep step, string inPath, string outPath)
    {
        Workbooks = workbooks;
        Step = step;
        InputFolder = inPath;
        OutputFolder = outPath;
    }

    public string GetParam(string key)
    {
        if (!Step.Params.ContainsKey(key))
            throw new ArgumentException($"Missing required parameter '{key}' for action '{Step.Action}'");

        object value = Step.Params[key];

        if (value is JsonElement element)
            return element.ToString();

        return value.ToString() ?? string.Empty;
    }

    public string? GetOptionalParam(string key)
    {
        if (!Step.Params.ContainsKey(key)) return null;

        object value = Step.Params[key];
        if (value is System.Text.Json.JsonElement element) return element.ToString();
        return value.ToString();
    }
}