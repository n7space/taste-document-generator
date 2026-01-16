using Avalonia;
using System;
using System.Collections.Generic;
using System.IO;
using CommandLine;

namespace TasteDocumentGenerator;

class Program
{

    [Verb("generate", HelpText = "Generate document in CLI mode")]
    public class GenerateOptions
    {
        [Option('t', "template-path", Required = true, HelpText = "Input template file path")]
        public string? TemplatePath { get; set; }

        [Option('i', "interface-view", Required = false, HelpText = "Input Interface View file path", Default = "interfaceview.xml")]
        public string InterfaceView { get; set; } = "interfaceview.xml";

        [Option('d', "deployment-view", Required = false, HelpText = "Input Deployment View file path", Default = "deploymentview.dv.xml")]
        public string DeploymentView { get; set; } = "deploymentview.dv.xml";

        [Option('p', "opus2-model-path", Required = true, HelpText = "Input OPUS2 model file path")]
        public string? Opus2Model { get; set; }

        [Option('o', "output-path", Required = true, HelpText = "Output file path")]
        public string? OutputPath { get; set; }

        [Option("target", Required = false, HelpText = "Target system", Default = "ASW")]
        public string Target { get; set; } = "ASW";

        [Option("template-directory", Required = false, HelpText = "Template directory path", Default = "")]
        public string TemplateDirectory { get; set; } = "";

        [Option("template-processor", Required = false, HelpText = "Template processor binary to execute [to support testing]", Default = "template-processor")]
        public string TemplateProcessor { get; set; } = "template-processor";

        [Option("system-object-exporter", Required = false, HelpText = "System Object Exporter binary path", Default = Orchestrator.DefaultSystemObjectExporterBinary)]
        public string SystemObjectExporter { get; set; } = Orchestrator.DefaultSystemObjectExporterBinary;

        [Option("system-object-type", Required = false, HelpText = "System Object Type names (comma separated)", Separator = ',')]
        public IEnumerable<string>? SystemObjectTypes { get; set; }

        [Option("tag", Required = false, HelpText = "Tag for template hooks (e.g., 'TDG:')", Default = "TDG:")]
        public string Tag { get; set; } = "TDG:";
    }

    [Verb("gui", HelpText = "Launch GUI")]
    public class GuiOptions
    {
        [Option('c', "configuration-path", Required = false, HelpText = "Configuration file path")]
        public string? OptionsPath { get; set; }
    }

    // Initialization code. Don't use any Avalonia, third-party APIs or any
    // SynchronizationContext-reliant code before AppMain is called: things aren't initialized
    // yet and stuff might break.
    [STAThread]
    public static void Main(string[] args)
    {
        Parser.Default.ParseArguments<GenerateOptions, GuiOptions>(args).WithParsed<GuiOptions>(o =>
        {
            Console.WriteLine("Launching GUI...");
            if (!string.IsNullOrEmpty(o.OptionsPath))
            {
                // Provide the configuration path to the rest of the app via environment variable
                Environment.SetEnvironmentVariable("TDG_SETTINGS_PATH", o.OptionsPath);
            }
            BuildAvaloniaApp()
                .StartWithClassicDesktopLifetime(args);
        }).WithParsed<GenerateOptions>(o =>
        {
            Console.WriteLine($"Generating document from {o.TemplatePath}");
            try
            {
                GenerateDocumentCli(o).Wait();
            }
            catch (Exception e)
            {
                Console.Error.WriteLine($"Error: {e.Message}");
                Environment.Exit(1);
            }
        });
    }

    private static async System.Threading.Tasks.Task GenerateDocumentCli(GenerateOptions options)
    {
        var orchestrator = new Orchestrator(new DocumentAssembler());
        var parameters = new Orchestrator.Parameters
        {
            TemplatePath = options.TemplatePath,
            InterfaceViewPath = options.InterfaceView,
            DeploymentViewPath = options.DeploymentView,
            Opus2ModelPath = options.Opus2Model,
            OutputPath = options.OutputPath,
            Target = options.Target,
            TemplateDirectory = options.TemplateDirectory,
            TemplateProcessorBinary = options.TemplateProcessor,
            SystemObjectExporterBinary = options.SystemObjectExporter,
            SystemObjectTypes = options.SystemObjectTypes,
            Tag = options.Tag
        };

        await orchestrator.GenerateAsync(parameters);

        Console.WriteLine($"Document {options.OutputPath} has been successfully created");
    }

    // Avalonia configuration, don't remove; also used by visual designer.
    public static AppBuilder BuildAvaloniaApp()
        => AppBuilder.Configure<App>()
            .UsePlatformDetect()
            .WithInterFont()
            .LogToTrace();
}
