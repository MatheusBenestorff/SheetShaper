using System.Collections.Generic;
using SheetShaper.Core;

namespace SheetShaper.Tests.Helpers;

public class JobBuilder
{
    private readonly SheetJob _job;

    public JobBuilder(string jobName)
    {
        _job = new SheetJob { JobName = jobName, Steps = new List<PipelineStep>() };
    }

    public JobBuilder AddLoadStep(string id, string file, string alias)
    {
        _job.Steps.Add(new PipelineStep 
        { 
            StepId = id, Action = "LoadSource", 
            Params = new Dictionary<string, object> { { "file", file }, { "alias", alias } } 
        });
        return this;
    }

    public JobBuilder AddSaveStep(string id, string fileName, string sourceAlias)
    {
        _job.Steps.Add(new PipelineStep 
        { 
            StepId = id, Action = "SaveFile", 
            Params = new Dictionary<string, object> { { "fileName", fileName }, { "sourceAlias", sourceAlias } } 
        });
        return this;
    }

    public JobBuilder AddMapStep(string id, string source, string target, object rules)
    {
        string jsonRules = System.Text.Json.JsonSerializer.Serialize(rules);
        _job.Steps.Add(new PipelineStep 
        { 
            StepId = id, Action = "MapColumns", 
            Params = new Dictionary<string, object> { { "sourceSheet", source }, { "targetSheet", target }, { "mappings", jsonRules } } 
        });
        return this;
    }

    public JobBuilder AddFilterStep(string id, string source, string target, string col, string op, object val)
    {
        _job.Steps.Add(new PipelineStep 
        { 
            StepId = id, Action = "FilterRows", 
            Params = new Dictionary<string, object> { 
                { "sourceSheet", source }, { "targetSheet", target }, 
                { "column", col }, { "operator", op }, { "value", val } 
            } 
        });
        return this;
    }

    public JobBuilder AddPivotStep(string id, string source, string target, string groupCol, string keyCol, string valCol)
    {
        _job.Steps.Add(new PipelineStep 
        { 
            StepId = id, Action = "PivotSheet", 
            Params = new Dictionary<string, object> { 
                { "sourceSheet", source }, 
                { "targetSheet", target }, 
                { "groupByColumn", groupCol }, 
                { "pivotKeyColumn", keyCol }, 
                { "pivotValueColumn", valCol } 
            } 
        });
        return this;
    }

    public SheetJob Build() => _job;
}