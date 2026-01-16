using System;
using System.Linq;
using System.Threading.Tasks;
using Avalonia;
using Avalonia.Platform.Storage;
using Avalonia.Controls.ApplicationLifetimes;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using System.IO;
using System.Xml.Serialization;
using MsBox.Avalonia;
using MsBox.Avalonia.Enums;

namespace TasteDocumentGenerator;

public partial class MainWindowViewModel : ObservableObject
{
    public class Settings
    {
        public string InputInterfaceViewPath { get; set; } = "interfaceview.xml";
        public string InputDeploymentViewPath { get; set; } = "deploymentview.dv.xml";
        public string InputOpus2ModelPath { get; set; } = "opus2_model.xml";
        public string InputTemplatePath { get; set; } = "sdd-template.docx";
        public string OutputFilePath { get; set; } = "sdd.docx";
        public string InputTemplateDirectoryPath { get; set; } = "";
        public string Target { get; set; } = "ASW";
        public bool DoOpenDocument { get; set; } = false;
        public string SystemObjectTypes { get; set; } = string.Join(", ", Orchestrator.DefaultSystemObjectTypes);
        public string Tag { get; set; } = "TDG:";
    }

    private static string GetSettingsFilePath()
    {
        var env = Environment.GetEnvironmentVariable("TDG_SETTINGS_PATH");
        if (!string.IsNullOrEmpty(env))
            return env!;
        return "taste-document-generator-settings.xml";
    }

    private void LoadSettings()
    {
        try
        {
            var settingsPath = GetSettingsFilePath();
            if (File.Exists(settingsPath))
            {
                var serializer = new XmlSerializer(typeof(Settings));
                using var stream = File.OpenRead(settingsPath);
                if (stream.Length > 0)
                {
                    var settings = (Settings?)serializer.Deserialize(stream);
                    if (settings != null)
                    {
                        InputInterfaceViewPath = settings.InputInterfaceViewPath;
                        InputDeploymentViewPath = settings.InputDeploymentViewPath;
                        InputOpus2ModelPath = settings.InputOpus2ModelPath;
                        InputTemplatePath = settings.InputTemplatePath;
                        OutputFilePath = settings.OutputFilePath;
                        Target = settings.Target;
                        InputTemplateDirectoryPath = settings.InputTemplateDirectoryPath;
                        DoOpenDocument = settings.DoOpenDocument;
                        SystemObjectTypesText = string.IsNullOrWhiteSpace(settings.SystemObjectTypes)
                            ? string.Join(", ", Orchestrator.DefaultSystemObjectTypes)
                            : settings.SystemObjectTypes;
                        TemplateTag = string.IsNullOrWhiteSpace(settings.Tag) ? "TDG:" : settings.Tag;
                    }
                }
            }
        }
        catch
        {
            // If loading fails, use default values
        }
    }

    private void SaveSettings()
    {
        try
        {
            var settings = new Settings
            {
                InputInterfaceViewPath = InputInterfaceViewPath,
                InputDeploymentViewPath = InputDeploymentViewPath,
                InputOpus2ModelPath = InputOpus2ModelPath,
                InputTemplatePath = InputTemplatePath,
                OutputFilePath = OutputFilePath,
                Target = Target,
                InputTemplateDirectoryPath = InputTemplateDirectoryPath,
                DoOpenDocument = DoOpenDocument,
                SystemObjectTypes = SystemObjectTypesText,
                Tag = TemplateTag
            };
            var serializer = new XmlSerializer(typeof(Settings));
            using var stream = File.Create(GetSettingsFilePath());
            serializer.Serialize(stream, settings);
        }
        catch (Exception e)
        {
            Console.Error.WriteLine($"Failed to save settings: {e}");

        }
    }

    public MainWindowViewModel()
    {
        LoadSettings();
        PropertyChanged += (s, e) => SaveSettings();
    }

    [ObservableProperty]
    private string _inputInterfaceViewPath = "interfaceview.xml";

    [ObservableProperty]
    private string _inputDeploymentViewPath = "deploymentview.dv.xml";

    [ObservableProperty]
    private string _inputOpus2ModelPath = "opus2_model.xml";

    [ObservableProperty]
    private string _inputTemplateDirectoryPath = "";

    [ObservableProperty]
    private string _target = "ASW";

    [ObservableProperty]
    private string _inputTemplatePath = "sdd-template.docx";

    [ObservableProperty]
    private string _outputFilePath = "sdd.docx";

    [ObservableProperty]
    private bool _doOpenDocument = false;

    [ObservableProperty]
    private string _systemObjectTypesText = string.Join(", ", Orchestrator.DefaultSystemObjectTypes);

    [ObservableProperty]
    private string _templateTag = "TDG:";

    private IStorageProvider GetStorageProvider()
    {
        if (Application.Current?.ApplicationLifetime is not IClassicDesktopStyleApplicationLifetime desktop ||
            desktop.MainWindow?.StorageProvider is not { } storageProvider)
        {
            throw new NullReferenceException("Missing StorageProvider instance.");
        }
        return storageProvider;
    }

    [RelayCommand]
    private async Task SelectInputTemplatePathAsync()
    {
        var storageProvider = GetStorageProvider();

        var file = await storageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
        {
            Title = "Select input template path",
            SuggestedFileType = new FilePickerFileType("docx"),
            SuggestedFileName = InputTemplatePath,
            FileTypeFilter = new[]
            {
                        new FilePickerFileType("Word Documents")
                        {
                            Patterns = new[] { "*.docx" }
                        }
                    }
        });

        if (file != null && file.Count > 0)
        {
            InputTemplatePath = file[0].Path.LocalPath;
        }
        await Task.CompletedTask;
    }

    [RelayCommand]
    private async Task SelectOutputFilePathAsync()
    {
        var storageProvider = GetStorageProvider();

        var file = await storageProvider.SaveFilePickerAsync(new FilePickerSaveOptions
        {
            Title = "Select output document path",
            DefaultExtension = "docx",
            SuggestedFileName = OutputFilePath,
            FileTypeChoices = new[]
            {
                        new FilePickerFileType("Word Documents")
                        {
                            Patterns = new[] { "*.docx" }
                        }
                    }
        });

        if (file != null)
        {
            OutputFilePath = file.Path.LocalPath;
        }
        await Task.CompletedTask;
    }

    private void OpenDocument(string path)
    {
        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
        {
            FileName = path,
            UseShellExecute = true
        });
    }

    [RelayCommand]
    private async Task GenerateDocumentAsync()
    {
        try
        {
            var orchestrator = new Orchestrator(new DocumentAssembler());
            var parameters = new Orchestrator.Parameters
            {
                TemplatePath = InputTemplatePath,
                InterfaceViewPath = InputInterfaceViewPath,
                DeploymentViewPath = InputDeploymentViewPath,
                Opus2ModelPath = InputOpus2ModelPath,
                OutputPath = OutputFilePath,
                Target = Target,
                TemplateDirectory = InputTemplateDirectoryPath,
                SystemObjectTypes = ParseSystemObjectTypes(SystemObjectTypesText),
                Tag = TemplateTag
            }; ;

            await orchestrator.GenerateAsync(parameters);
            var messageBox = MessageBoxManager.GetMessageBoxStandard(
                "Success",
                $"Document {OutputFilePath} has been successfully created",
                ButtonEnum.Ok,
                Icon.Info);
            await messageBox.ShowAsync();
            if (DoOpenDocument)
            {
                OpenDocument(OutputFilePath);
            }
        }
        catch (Exception e)
        {
            var messageBox = MessageBoxManager.GetMessageBoxStandard(
                "Error",
                $"An error occurred while generating the document:\n\n{e.Message}",
                ButtonEnum.Ok,
                Icon.Error);
            await messageBox.ShowAsync();
        }

    }

    private static string[] ParseSystemObjectTypes(string? text)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return Array.Empty<string>();
        }

        return text
            .Split(',', StringSplitOptions.RemoveEmptyEntries)
            .Select(part => part.Trim())
            .Where(part => !string.IsNullOrWhiteSpace(part))
            .ToArray();
    }
    [RelayCommand]
    private async Task ShowAboutAsync()
    {
        var messageBox = MessageBoxManager.GetMessageBoxStandard(
            "About",
            $"TASTE Document Generator\nCreated by N7 Space\n within the scope of Model Based Execution Platform Project\nESA Contract 4000146882/24/NL/KK",
            ButtonEnum.Ok,
            Icon.Info);
        await messageBox.ShowAsync();
    }
}
