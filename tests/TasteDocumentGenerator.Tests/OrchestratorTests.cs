using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace TasteDocumentGenerator.Tests;

public class OrchestratorTests
{
    [Fact]
    public async Task GenerateAsync_PassesExporterOutputsToDocumentAssembler()
    {
        var tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
        Directory.CreateDirectory(tempDir);

        var templatePath = Path.Combine(tempDir, "template.docx");
        var interfaceViewPath = Path.Combine(tempDir, "iv.xml");
        var deploymentViewPath = Path.Combine(tempDir, "dv.xml");
        var opusPath = Path.Combine(tempDir, "opus.xml");
        var outputPath = Path.Combine(tempDir, "output.docx");

        File.WriteAllText(templatePath, "template");
        File.WriteAllText(interfaceViewPath, "iv");
        File.WriteAllText(deploymentViewPath, "dv");
        File.WriteAllText(opusPath, "opus");

        try
        {
            var runner = new TestProcessRunner();
            var assembler = new TestDocumentAssembler();
            var orchestrator = new Orchestrator(assembler, runner);
            var parameters = new Orchestrator.Parameters
            {
                TemplatePath = templatePath,
                InterfaceViewPath = interfaceViewPath,
                DeploymentViewPath = deploymentViewPath,
                Opus2ModelPath = opusPath,
                OutputPath = outputPath,
                Target = "CubeSat",
                TemplateDirectory = string.Empty,
                SystemObjectExporterBinary = runner.ExpectedBinary,
                SystemObjectTypes = new[] { "On-board memory", "File System" }
            };

            await orchestrator.GenerateAsync(parameters);

            Assert.Equal(2, runner.Calls.Count);
            Assert.NotNull(assembler.CapturedContext);
            Assert.Equal(2, assembler.CapturedContext!.SystemObjectCsvFiles.Count);

            var csvNames = assembler.CapturedContext.SystemObjectCsvFiles.Select(Path.GetFileName).ToArray();
            Assert.Contains("on_board_memory.csv", csvNames);
            Assert.Contains("file_system.csv", csvNames);
        }
        finally
        {
            Directory.Delete(tempDir, true);
        }
    }

    [Fact]
    public async Task GenerateAsync_ThrowsWhenExporterFails()
    {
        var tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
        Directory.CreateDirectory(tempDir);
        var templatePath = Path.Combine(tempDir, "template.docx");
        var interfaceViewPath = Path.Combine(tempDir, "iv.xml");
        var deploymentViewPath = Path.Combine(tempDir, "dv.xml");
        var opusPath = Path.Combine(tempDir, "opus.xml");
        var outputPath = Path.Combine(tempDir, "output.docx");
        File.WriteAllText(templatePath, "template");
        File.WriteAllText(interfaceViewPath, "iv");
        File.WriteAllText(deploymentViewPath, "dv");
        File.WriteAllText(opusPath, "opus");

        try
        {
            var runner = new TestProcessRunner { ShouldFail = true };
            var assembler = new TestDocumentAssembler();
            var orchestrator = new Orchestrator(assembler, runner);
            var parameters = new Orchestrator.Parameters
            {
                TemplatePath = templatePath,
                InterfaceViewPath = interfaceViewPath,
                DeploymentViewPath = deploymentViewPath,
                Opus2ModelPath = opusPath,
                OutputPath = outputPath,
                Target = "CubeSat",
                SystemObjectExporterBinary = runner.ExpectedBinary,
                SystemObjectTypes = new[] { "On-board memory" }
            };

            await Assert.ThrowsAsync<InvalidOperationException>(() => orchestrator.GenerateAsync(parameters));
        }
        finally
        {
            Directory.Delete(tempDir, true);
        }
    }

    private sealed class TestDocumentAssembler : IDocumentAssembler
    {
        public DocumentAssembler.Context? CapturedContext { get; private set; }

        public Task ProcessTemplate(DocumentAssembler.Context context, string inputTemplatePath, string outputDocumentPath)
        {
            CapturedContext = context;
            return Task.CompletedTask;
        }
    }

    private sealed class TestProcessRunner : IProcessRunner
    {
        public string ExpectedBinary { get; } = "mock-exporter";
        public bool ShouldFail { get; set; }
        public List<(string FileName, IReadOnlyList<string> Arguments)> Calls { get; } = new();

        public Task<ProcessResult> RunAsync(ProcessStartInfo processStartInfo, CancellationToken cancellationToken = default)
        {
            Calls.Add((processStartInfo.FileName, processStartInfo.ArgumentList.ToList()));
            var outputPath = GetArgumentValue(processStartInfo.ArgumentList, "--output");
            if (!string.IsNullOrEmpty(outputPath))
            {
                Directory.CreateDirectory(Path.GetDirectoryName(outputPath)!);
                File.WriteAllText(outputPath, "name,value\nexample,data");
            }

            if (ShouldFail)
            {
                return Task.FromResult(new ProcessResult(1, string.Empty, "failed"));
            }

            return Task.FromResult(new ProcessResult(0, string.Empty, string.Empty));
        }

        private static string? GetArgumentValue(IList<string> arguments, string key)
        {
            for (var i = 0; i < arguments.Count - 1; i++)
            {
                if (string.Equals(arguments[i], key, StringComparison.Ordinal))
                {
                    return arguments[i + 1];
                }
            }

            return null;
        }
    }
}