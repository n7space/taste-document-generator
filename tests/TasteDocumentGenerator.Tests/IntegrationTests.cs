using System.Diagnostics;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

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
        Assert.Contains("--system-object-exporter", output, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("--system-object-type", output, StringComparison.OrdinalIgnoreCase);
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
            var opusPath = Path.Combine(tempDir, "dummy.xml");
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

            var mockExporter = Path.Combine(repoRoot, "data", "mock-exporter.sh");
            Assert.True(File.Exists(mockExporter), $"Mock exporter not found: {mockExporter}");
            var exporterLog = Path.Combine(repoRoot, "data", "mock-exporter.output");
            if (File.Exists(exporterLog))
            {
                File.Delete(exporterLog);
            }

            var psi = new ProcessStartInfo
            {
                FileName = "dotnet",
                Arguments = $"run --project \"src/TasteDocumentGenerator/TasteDocumentGenerator.csproj\" -- generate -t \"{templatePath}\" -i \"{ivPath}\" -d \"{dvPath}\" -p \"{opusPath}\" -o \"{outputPath}\" --target CubeSat --template-processor \"{mockProcessor}\" --system-object-exporter \"{mockExporter}\"",
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
    public async Task CliInterface_NoTargetSkipsCsvExtraction()
    {
        var repoRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../.."));
        var tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
        Directory.CreateDirectory(tempDir);
        var logPath = Path.Combine(repoRoot, "data", "mock-processor.output");
        if (File.Exists(logPath))
        {
            File.Delete(logPath);
        }

        var exporterLogPath = Path.Combine(repoRoot, "data", "mock-exporter.output");
        if (File.Exists(exporterLogPath))
        {
            File.Delete(exporterLogPath);
        }

        try
        {
            var ivPath = Path.Combine(tempDir, "dummy.iv");
            var dvPath = Path.Combine(tempDir, "dummy.dv");
            var opusPath = Path.Combine(tempDir, "dummy.xml");
            File.WriteAllText(ivPath, "<InterfaceView />");
            File.WriteAllText(dvPath, "<DeploymentView />");
            File.WriteAllText(opusPath, "<Opus />");

            var templatePath = Path.Combine(repoRoot, "data", "test_in_simple.docx");
            Assert.True(File.Exists(templatePath), $"Template file not found: {templatePath}");

            var outputDir = Path.Combine(tempDir, "output");
            Directory.CreateDirectory(outputDir);
            var outputPath = Path.Combine(outputDir, "no-target.docx");

            var mockProcessor = Path.Combine(repoRoot, "data", "mock-processor.sh");
            Assert.True(File.Exists(mockProcessor), $"Mock processor not found: {mockProcessor}");

            var mockExporter = Path.Combine(repoRoot, "data", "mock-exporter.sh");
            Assert.True(File.Exists(mockExporter), $"Mock exporter not found: {mockExporter}");

            var psi = new ProcessStartInfo
            {
                FileName = "dotnet",
                Arguments = $"run --project \"src/TasteDocumentGenerator/TasteDocumentGenerator.csproj\" -- generate -t \"{templatePath}\" -i \"{ivPath}\" -d \"{dvPath}\" -p \"{opusPath}\" -o \"{outputPath}\" --template-processor \"{mockProcessor}\" --system-object-exporter \"{mockExporter}\"",
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                WorkingDirectory = repoRoot
            };

            using var process = Process.Start(psi) ?? throw new InvalidOperationException("Failed to start dotnet run process");
            var stdoutTask = process.StandardOutput.ReadToEndAsync();
            var stderrTask = process.StandardError.ReadToEndAsync();
            await process.WaitForExitAsync();
            await stdoutTask;
            await stderrTask;

            Assert.Equal(0, process.ExitCode);
            Assert.True(File.Exists(outputPath), "Output file should be created");

            // Verify mock processor was called and that it was NOT passed a TARGET value
            Assert.True(File.Exists(logPath), "Mock processor log should exist");
            var logContents = await File.ReadAllTextAsync(logPath);
            Assert.DoesNotContain("TARGET", logContents);

            // Verify CSV extraction was skipped: exporter should not have been invoked
            Assert.False(File.Exists(exporterLogPath), "Mock exporter log should not exist when no target provided");
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

            try
            {
                if (File.Exists(exporterLogPath))
                {
                    File.Delete(exporterLogPath);
                }
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
            var opusPath = Path.Combine(tempDir, "dummy.xml");
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

            var mockExporter = Path.Combine(repoRoot, "data", "mock-exporter.sh");
            Assert.True(File.Exists(mockExporter), $"Mock exporter not found: {mockExporter}");
            var exporterLogPath = Path.Combine(repoRoot, "data", "mock-exporter.output");
            if (File.Exists(exporterLogPath))
            {
                File.Delete(exporterLogPath);
            }

            var psi = new ProcessStartInfo
            {
                FileName = "dotnet",
                Arguments = $"run --project \"src/TasteDocumentGenerator/TasteDocumentGenerator.csproj\" -- generate -t \"{templatePath}\" -i \"{ivPath}\" -d \"{dvPath}\" -p \"{opusPath}\" -o \"{outputPath}\" --target CubeSat --template-processor \"{mockProcessor}\" --system-object-exporter \"{mockExporter}\"",
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

            Assert.True(File.Exists(exporterLogPath), "Mock exporter log should exist");
            var exporterLogContents = await File.ReadAllTextAsync(exporterLogPath);
            Assert.Contains("--system-object-type", exporterLogContents, StringComparison.Ordinal);

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

            try
            {
                if (File.Exists(Path.Combine(repoRoot, "data", "mock-exporter.output")))
                {
                    File.Delete(Path.Combine(repoRoot, "data", "mock-exporter.output"));
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

    [Fact]
    public async Task CliInterface_MergesDocumentWithImages()
    {
        /*
        # Test specification (DO NOT REMOVE)
        TasteDocumentGenerator is executed using dotnet run, in generate mode, with the following arguments:
        -InterfaceView is set to dummy.iv
        -DeploymentView is set to dummy.dv
        -Target is set to TestTarget
        -OPUS2 Model is set to dummy.xml
        -Input file is a template containing a document command referencing images.docx
        -Output file is output/with-images.docx
        -Template processor is set to data/mock-processor.sh 
        The final produced output contains the images.docx content merged inside, including images.
        All image parts are properly copied and image references are valid.
        */
        var repoRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../.."));
        var tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
        Directory.CreateDirectory(tempDir);

        try
        {
            var ivPath = Path.Combine(tempDir, "dummy.iv");
            var dvPath = Path.Combine(tempDir, "dummy.dv");
            var opusPath = Path.Combine(tempDir, "dummy.xml");
            File.WriteAllText(ivPath, "<InterfaceView />");
            File.WriteAllText(dvPath, "<DeploymentView />");
            File.WriteAllText(opusPath, "<Opus />");

            // Copy images.docx to the temp directory
            var imagesDocPath = Path.Combine(repoRoot, "data", "images.docx");
            Assert.True(File.Exists(imagesDocPath), $"Images document not found: {imagesDocPath}");
            var localImagesPath = Path.Combine(tempDir, "images.docx");
            File.Copy(imagesDocPath, localImagesPath);

            // Create a template with document command referencing images.docx
            var templatePath = Path.Combine(tempDir, "template-with-doc-cmd.docx");
            using (var templateDoc = WordprocessingDocument.Create(templatePath, WordprocessingDocumentType.Document))
            {
                var mainPart = templateDoc.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());
                var body = mainPart.Document.Body;
                body!.Append(new Paragraph(new Run(new Text("Before merged document"))));
                body.Append(new Paragraph(new Run(new Text("<TDG: document images.docx />"))));
                body.Append(new Paragraph(new Run(new Text("After merged document"))));
            }

            var outputDir = Path.Combine(tempDir, "output");
            Directory.CreateDirectory(outputDir);
            var outputPath = Path.Combine(outputDir, "with-images.docx");

            var mockProcessor = Path.Combine(repoRoot, "data", "mock-processor.sh");
            Assert.True(File.Exists(mockProcessor), $"Mock processor not found: {mockProcessor}");

            var mockExporter = Path.Combine(repoRoot, "data", "mock-exporter.sh");
            Assert.True(File.Exists(mockExporter), $"Mock exporter not found: {mockExporter}");

            var psi = new ProcessStartInfo
            {
                FileName = "dotnet",
                Arguments = $"run --project \"src/TasteDocumentGenerator/TasteDocumentGenerator.csproj\" -- generate -t \"{templatePath}\" -i \"{ivPath}\" -d \"{dvPath}\" -p \"{opusPath}\" -o \"{outputPath}\" --target TestTarget --template-directory \"{tempDir}\" --template-processor \"{mockProcessor}\" --system-object-exporter \"{mockExporter}\"",
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

            if (process.ExitCode != 0)
            {
                throw new Exception($"Process failed with exit code {process.ExitCode}. STDOUT: {stdout}. STDERR: {stderr}");
            }

            Assert.Equal(0, process.ExitCode);
            Assert.True(File.Exists(outputPath), "Output file should be created");

            // Verify the output document
            using (var outputDoc = WordprocessingDocument.Open(outputPath, false))
            {
                var body = outputDoc.MainDocumentPart!.Document!.Body!;
                var outputText = ReadDocumentText(outputPath);

                // Verify template text
                Assert.Contains("Before merged document", outputText);
                Assert.Contains("After merged document", outputText);

                // Verify images were copied
                var imageParts = outputDoc.MainDocumentPart.ImageParts.ToList();
                Assert.NotEmpty(imageParts);
                // Reference document contains documentation of 2 SDL processes, each with 3 images
                Assert.Equal(6, imageParts.Count);
                // Verify all image references are valid
                var blips = body.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().ToList();
                if (blips.Any())
                {
                    foreach (var blip in blips)
                    {
                        var embedId = blip.Embed?.Value;
                        Assert.NotNull(embedId);
                        var part = outputDoc.MainDocumentPart.GetPartById(embedId!);
                        Assert.NotNull(part);
                    }
                }
            }

            // Verify the source images.docx has images to merge
            using (var sourceDoc = WordprocessingDocument.Open(localImagesPath, false))
            {
                var sourceImageParts = sourceDoc.MainDocumentPart!.ImageParts.ToList();
                Assert.NotEmpty(sourceImageParts);
            }
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
        }
    }
}
