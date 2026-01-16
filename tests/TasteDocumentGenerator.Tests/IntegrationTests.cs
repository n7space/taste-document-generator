using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace TasteDocumentGenerator.Tests;

public class IntegrationTests
{
    [Fact]
    public async Task CliInterface_HelpIsPrintedForGenerate()
    {
        /**
        # Test specification (DO NOT REMOVE)
        TasteDocumentGenerator is executed using dotnet run, with arguments generate --help
        The application output is the contents of the built-in help, which includes at least the following items:
        --template-path
        --interface-view
        --deployment-view
        --target
        **/
        var repoRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../.."));
        var psi = new ProcessStartInfo
        {
            FileName = "dotnet",
            Arguments = "run --project \"src/TasteDocumentGenerator/TasteDocumentGenerator.csproj\" -- generate --help",
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            WorkingDirectory = repoRoot
        };

        using var process = Process.Start(psi) ?? throw new InvalidOperationException("Failed to start dotnet run process");
        var stdoutTask = process.StandardOutput.ReadToEndAsync();
        var stderrTask = process.StandardError.ReadToEndAsync();
        await process.WaitForExitAsync();
        var stdout = await stdoutTask;
        var stderr = await stderrTask;

        Assert.Equal(0, process.ExitCode);

        var output = stdout + stderr;
        Assert.Contains("--template-path", output, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("--interface-view", output, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("--deployment-view", output, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("--target", output, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task CliInterface_CreatesTargetFile()
    {
        /*
        # Test specification (DO NOT REMOVE)
        TasteDocumentGenerator is executed using dotnet run, in generate mode, with the following arguments:
        -InterfaceView is set to dummy.iv
        -DeploymentView is set to dummy.dv
        -Target is set to CubeSat
        -OPUS2 Model is set to dummy.opus
        -Input file is data/test_in_empty.docx
        -Output file is output/empty.docx
        -Template processor is set to data/mock-processor.sh 
        The produced output is the same as the input
        */
        var repoRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../.."));
        var tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
        Directory.CreateDirectory(tempDir);
        try
        {
            var ivPath = Path.Combine(tempDir, "dummy.iv");
            var dvPath = Path.Combine(tempDir, "dummy.dv");
            var opusPath = Path.Combine(tempDir, "dummy.opus");
            File.WriteAllText(ivPath, "<InterfaceView />");
            File.WriteAllText(dvPath, "<DeploymentView />");
            File.WriteAllText(opusPath, "<Opus />");

            var templatePath = Path.Combine(repoRoot, "data", "test_in_empty.docx");
            Assert.True(File.Exists(templatePath), $"Template file not found: {templatePath}");

            var outputDir = Path.Combine(tempDir, "output");
            Directory.CreateDirectory(outputDir);
            var outputPath = Path.Combine(outputDir, "empty.docx");

            var mockProcessor = Path.Combine(repoRoot, "data", "mock-processor.sh");
            Assert.True(File.Exists(mockProcessor), $"Mock processor not found: {mockProcessor}");

            var psi = new ProcessStartInfo
            {
                FileName = "dotnet",
                Arguments = $"run --project \"src/TasteDocumentGenerator/TasteDocumentGenerator.csproj\" -- generate -t \"{templatePath}\" -i \"{ivPath}\" -d \"{dvPath}\" -p \"{opusPath}\" -o \"{outputPath}\" --target CubeSat --template-processor \"{mockProcessor}\"",
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                WorkingDirectory = repoRoot
            };

            using var process = Process.Start(psi) ?? throw new InvalidOperationException("Failed to start dotnet run process");
            var stdoutTask = process.StandardOutput.ReadToEndAsync();
            var stderrTask = process.StandardError.ReadToEndAsync();
            await process.WaitForExitAsync();
            var stdout = await stdoutTask;
            var stderr = await stderrTask;

            Assert.Equal(0, process.ExitCode);
            Assert.True(File.Exists(outputPath), "Output file should be created");

            var inputText = ReadDocumentText(templatePath);
            var outputText = ReadDocumentText(outputPath);
            Assert.Equal(inputText, outputText);
        }
        finally
        {
            try
            {
                Directory.Delete(tempDir, true);
            }
            catch
            {
            }
        }
    }

    [Fact]
    public async Task CliInterface_ExecutesTemplate()
    {
        /*
        # Test specification (DO NOT REMOVE)
        TasteDocumentGenerator is executed using dotnet run, in generate mode, with the following arguments:
        -InterfaceView is set to dummy.iv
        -DeploymentView is set to dummy.dv
        -Target is set to CubeSat
        -OPUS2 Model is set to dummy.opus
        -Input file is data/test_in_simple.docx
        -Output file is output/simple.docx
        -Template processor is set to data/mock-processory.sh 
        Mock template processor was called with the proper InterfaceView, DeploymentView, input and output parameters.
        The final produced output contains the dummy test_in_tmplt.docx merged inside.
        */
        var repoRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../.."));
        var tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
        Directory.CreateDirectory(tempDir);
        var logPath = Path.Combine(repoRoot, "data", "mock-processor.output");
        if (File.Exists(logPath))
        {
            File.Delete(logPath);
        }

        try
        {
            var ivPath = Path.Combine(tempDir, "dummy.iv");
            var dvPath = Path.Combine(tempDir, "dummy.dv");
            var opusPath = Path.Combine(tempDir, "dummy.opus");
            File.WriteAllText(ivPath, "<InterfaceView />");
            File.WriteAllText(dvPath, "<DeploymentView />");
            File.WriteAllText(opusPath, "<Opus />");

            var templatePath = Path.Combine(repoRoot, "data", "test_in_simple.docx");
            Assert.True(File.Exists(templatePath), $"Template file not found: {templatePath}");

            var outputDir = Path.Combine(tempDir, "output");
            Directory.CreateDirectory(outputDir);
            var outputPath = Path.Combine(outputDir, "simple.docx");

            var mockProcessor = Path.Combine(repoRoot, "data", "mock-processor.sh");
            Assert.True(File.Exists(mockProcessor), $"Mock processor not found: {mockProcessor}");

            var psi = new ProcessStartInfo
            {
                FileName = "dotnet",
                Arguments = $"run --project \"src/TasteDocumentGenerator/TasteDocumentGenerator.csproj\" -- generate -t \"{templatePath}\" -i \"{ivPath}\" -d \"{dvPath}\" -p \"{opusPath}\" -o \"{outputPath}\" --target CubeSat --template-processor \"{mockProcessor}\"",
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                WorkingDirectory = repoRoot
            };

            using var process = Process.Start(psi) ?? throw new InvalidOperationException("Failed to start dotnet run process");
            var stdoutTask = process.StandardOutput.ReadToEndAsync();
            var stderrTask = process.StandardError.ReadToEndAsync();
            await process.WaitForExitAsync();
            var stdout = await stdoutTask;
            var stderr = await stderrTask;

            Assert.Equal(0, process.ExitCode);
            Assert.True(File.Exists(outputPath), "Output file should be created");

            Assert.True(File.Exists(logPath), "Mock processor log should exist");
            var logContents = await File.ReadAllTextAsync(logPath);
            Assert.Contains($"--iv {ivPath}", logContents, StringComparison.Ordinal);
            Assert.Contains($"--dv {dvPath}", logContents, StringComparison.Ordinal);
            Assert.Contains("-t test_in_tmplt.tmplt", logContents, StringComparison.Ordinal);
            Assert.Contains(" -o ", logContents, StringComparison.Ordinal);

            var outputText = ReadDocumentText(outputPath);
            var insertedText = ReadDocumentText(Path.Combine(repoRoot, "data", "test_in_tmplt.docx"));
            Assert.Contains("DOCUMENT BEGIN", outputText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("DOCUMENT END", outputText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains(insertedText, outputText, StringComparison.OrdinalIgnoreCase);
        }
        finally
        {
            try
            {
                if (Directory.Exists(tempDir))
                {
                    Directory.Delete(tempDir, true);
                }
            }
            catch
            {
            }

            try
            {
                if (File.Exists(logPath))
                {
                    File.Delete(logPath);
                }
            }
            catch
            {
            }
        }
    }

    private static string ReadDocumentText(string path)
    {
        using var doc = WordprocessingDocument.Open(path, false);
        var body = doc.MainDocumentPart?.Document?.Body;
        return body == null
            ? string.Empty
            : string.Concat(body.Descendants<Text>().Select(t => t.Text));
    }
}
