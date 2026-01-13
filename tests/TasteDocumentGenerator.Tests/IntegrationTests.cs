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
}
