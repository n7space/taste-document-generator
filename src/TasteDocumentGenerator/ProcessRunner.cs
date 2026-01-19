namespace TasteDocumentGenerator;

using System;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;

public interface IProcessRunner
{
    Task<ProcessResult> RunAsync(ProcessStartInfo processStartInfo, CancellationToken cancellationToken = default);
}

public sealed record ProcessResult(int ExitCode, string StandardOutput, string StandardError);

public class ProcessRunner : IProcessRunner
{
    public async Task<ProcessResult> RunAsync(ProcessStartInfo processStartInfo, CancellationToken cancellationToken = default)
    {
        using var process = Process.Start(processStartInfo) ?? throw new InvalidOperationException($"Failed to start process {processStartInfo.FileName}");
        var stdoutTask = process.StandardOutput.ReadToEndAsync();
        var stderrTask = process.StandardError.ReadToEndAsync();
        await process.WaitForExitAsync(cancellationToken);
        var stdout = await stdoutTask;
        var stderr = await stderrTask;
        return new ProcessResult(process.ExitCode, stdout, stderr);
    }
}