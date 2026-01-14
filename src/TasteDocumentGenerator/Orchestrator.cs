namespace TasteDocumentGenerator;

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

public sealed class Orchestrator
{
    public const string DefaultSystemObjectExporterBinary = "Opus2.SystemObjectCLIExporter";

    public static readonly IReadOnlyList<string> DefaultSystemObjectTypes = Array.AsReadOnly(new[]
    {
        "On-board memory",
        "On-board parameter",
        "File System",
        "Event definition",
        "Housekeeping parameter report structure"
    });

    private readonly IDocumentAssembler _documentAssembler;
    private readonly IProcessRunner _processRunner;

    public Orchestrator(IDocumentAssembler documentAssembler, IProcessRunner? processRunner = null)
    {
        _documentAssembler = documentAssembler ?? throw new ArgumentNullException(nameof(documentAssembler));
        _processRunner = processRunner ?? new ProcessRunner();
    }

    public sealed class Parameters
    {
        public string? TemplatePath { get; init; }
        public string? InterfaceViewPath { get; init; }
        public string? DeploymentViewPath { get; init; }
        public string? Opus2ModelPath { get; init; }
        public string? OutputPath { get; init; }
        public string Target { get; init; } = "ASW";
        public string TemplateDirectory { get; init; } = string.Empty;
        public string? TemplateProcessorBinary { get; init; }
        public string? SystemObjectExporterBinary { get; init; }
        public IEnumerable<string>? SystemObjectTypes { get; init; }
    }

    public async Task GenerateAsync(Parameters parameters, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(parameters);

        var templatePath = EnsureExistingFile(parameters.TemplatePath, nameof(parameters.TemplatePath));
        var interfaceViewPath = EnsureExistingFile(parameters.InterfaceViewPath, nameof(parameters.InterfaceViewPath));
        var deploymentViewPath = EnsureExistingFile(parameters.DeploymentViewPath, nameof(parameters.DeploymentViewPath));
        var opus2ModelPath = EnsureExistingFile(parameters.Opus2ModelPath, nameof(parameters.Opus2ModelPath));
        var outputPath = EnsureWritablePath(parameters.OutputPath, nameof(parameters.OutputPath));
        var target = EnsureNotEmpty(parameters.Target, nameof(parameters.Target));
        var templateDirectory = parameters.TemplateDirectory ?? string.Empty;
        var templateProcessorBinary = parameters.TemplateProcessorBinary;
        var exporterBinary = string.IsNullOrWhiteSpace(parameters.SystemObjectExporterBinary)
            ? DefaultSystemObjectExporterBinary
            : parameters.SystemObjectExporterBinary!;
        var systemObjectTypes = NormalizeSystemObjectTypes(parameters.SystemObjectTypes);

        var workingDirectory = Path.Combine(Path.GetTempPath(), $"tdg_{Guid.NewGuid():N}");
        var exporterOutputDirectory = Path.Combine(workingDirectory, "exports");
        var assemblerTempDirectory = Path.Combine(workingDirectory, "assembler");
        Directory.CreateDirectory(exporterOutputDirectory);
        Directory.CreateDirectory(assemblerTempDirectory);

        try
        {
            var csvFiles = await ExportSystemObjectDataAsync(
                exporterBinary,
                systemObjectTypes,
                opus2ModelPath,
                target,
                exporterOutputDirectory,
                cancellationToken).ConfigureAwait(false);

            var context = new DocumentAssembler.Context(
                interfaceViewPath,
                deploymentViewPath,
                target,
                templateDirectory,
                assemblerTempDirectory,
                templateProcessorBinary,
                csvFiles);

            await _documentAssembler.ProcessTemplate(context, templatePath, outputPath).ConfigureAwait(false);
        }
        finally
        {
            TryDeleteDirectory(workingDirectory);
        }
    }

    private static IReadOnlyList<string> NormalizeSystemObjectTypes(IEnumerable<string>? requestedTypes)
    {
        var normalized = requestedTypes?
            .Select(type => type?.Trim())
            .Where(type => !string.IsNullOrWhiteSpace(type))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        return normalized is { Count: > 0 } ? normalized : DefaultSystemObjectTypes;
    }

    private async Task<IReadOnlyList<string>> ExportSystemObjectDataAsync(
        string exporterBinary,
        IReadOnlyList<string> systemObjectTypes,
        string opus2ModelPath,
        string target,
        string outputDirectory,
        CancellationToken cancellationToken)
    {
        var csvFiles = new List<string>(systemObjectTypes.Count);
        foreach (var systemObjectType in systemObjectTypes)
        {
            var csvPath = Path.Combine(outputDirectory, BuildCsvFileName(systemObjectType));
            Directory.CreateDirectory(Path.GetDirectoryName(csvPath)!);
            var startInfo = new ProcessStartInfo
            {
                FileName = exporterBinary,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            startInfo.ArgumentList.Add("--model");
            startInfo.ArgumentList.Add(opus2ModelPath);
            startInfo.ArgumentList.Add("--deployment-target");
            startInfo.ArgumentList.Add(target);
            startInfo.ArgumentList.Add("--system-object-type");
            startInfo.ArgumentList.Add(systemObjectType);
            startInfo.ArgumentList.Add("--output");
            startInfo.ArgumentList.Add(csvPath);

            Debug.WriteLine($"Exporting {target}:{systemObjectType} to {csvPath}");

            var result = await _processRunner.RunAsync(startInfo, cancellationToken).ConfigureAwait(false);
            if (result.ExitCode != 0)
            {
                throw new InvalidOperationException($"System Object Exporter exited with code {result.ExitCode}: {result.StandardError}");
            }

            if (!File.Exists(csvPath))
            {
                throw new FileNotFoundException($"System Object Exporter did not create expected file {csvPath}", csvPath);
            }

            csvFiles.Add(csvPath);
        }

        return csvFiles;
    }

    private static string BuildCsvFileName(string systemObjectType)
    {
        var sanitized = new StringBuilder();
        var trimmed = systemObjectType?.Trim() ?? string.Empty;
        var fallback = "system_object";
        foreach (var c in trimmed)
        {
            if (char.IsLetterOrDigit(c))
            {
                sanitized.Append(char.ToLowerInvariant(c));
            }
            else
            {
                sanitized.Append('_');
            }
        }

        var fileName = sanitized.Length == 0 ? fallback : sanitized.ToString();
        return $"{fileName}.csv";
    }

    private static string EnsureExistingFile(string? path, string argumentName)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            throw new ArgumentException($"Parameter {argumentName} must be provided.", argumentName);
        }

        if (!File.Exists(path))
        {
            throw new FileNotFoundException($"File not found: {path}", path);
        }

        return path;
    }

    private static string EnsureWritablePath(string? path, string argumentName)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            throw new ArgumentException($"Parameter {argumentName} must be provided.", argumentName);
        }

        var directory = Path.GetDirectoryName(path);
        if (!string.IsNullOrWhiteSpace(directory))
        {
            Directory.CreateDirectory(directory);
        }

        return path;
    }

    private static string EnsureNotEmpty(string? value, string argumentName)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            throw new ArgumentException($"Parameter {argumentName} must be provided.", argumentName);
        }
        return value;
    }

    private static void TryDeleteDirectory(string directory)
    {
        try
        {
            if (Directory.Exists(directory))
            {
                Directory.Delete(directory, true);
            }
        }
        catch
        {
            // Swallow cleanup errors
        }
    }
}