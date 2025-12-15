using Avalonia;
using System;
using CommandLine;

namespace TasteDocumentGenerator;

class Program
{

    [Verb("generate", HelpText = "Generate document in CLI mode")]
    public class GenerateOptions
    {
        [Option('t', "template-path", Required = true, HelpText = "Input template file path")]
        public string? TemplatePath { get; set; }

        [Option('i', "interface-view", Required = false, HelpText = "Input Interface View file path")]
        public string? InterfaceView { get; set; }

        [Option('d', "deployment-view", Required = false, HelpText = "Input Deployment View file path")]
        public string? DeploymentView { get; set; }

        [Option('p', "opus2-model-path", Required = false, HelpText = "Input OPUS2 model file path")]
        public string? Opus2Model { get; set; }

        [Option('o', "output-path", Required = true, HelpText = "Output file path")]
        public string? OutputPath { get; set; }
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
            BuildAvaloniaApp()
                .StartWithClassicDesktopLifetime(args);
        }).WithParsed<GenerateOptions>(o =>
        {
            Console.WriteLine($"Generating document from {o.TemplatePath}");
        });
    }

    // Avalonia configuration, don't remove; also used by visual designer.
    public static AppBuilder BuildAvaloniaApp()
        => AppBuilder.Configure<App>()
            .UsePlatformDetect()
            .WithInterFont()
            .LogToTrace();
}
