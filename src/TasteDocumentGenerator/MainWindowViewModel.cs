using System;
using System.Threading.Tasks;
using System.Windows.Input;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Platform.Storage;
using Avalonia.Controls.ApplicationLifetimes;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using System.IO;
using System.Text.Json;
using System.Linq;

namespace TasteDocumentGenerator;

public partial class MainWindowViewModel : ObservableObject
{
    private class Settings
    {
        public string InputInterfaceViewPath { get; set; } = "interfaceview.xml";
        public string InputDeploymentViewPath { get; set; } = "deploymentview.dv.xml";
        public string InputOpus2ModelPath { get; set; } = "opus2_model.xml";
        public string InputTemplatePath { get; set; } = "sdd-template.docx";
        public string OutputFilePath { get; set; } = "sdd.docx";
        public string InputTemplateDirectoryPath { get; set; } = "";
        public string Target { get; set; } = "ASW";
    }

    private static string GetSettingsFilePath()
    {
        return "taste-document-generator-settings.json";
    }

    private void LoadSettings()
    {
        try
        {
            var settingsPath = GetSettingsFilePath();
            if (File.Exists(settingsPath))
            {
                var json = File.ReadAllText(settingsPath);
                var settings = JsonSerializer.Deserialize<Settings>(json);
                if (settings != null)
                {
                    InputInterfaceViewPath = settings.InputInterfaceViewPath;
                    InputDeploymentViewPath = settings.InputDeploymentViewPath;
                    InputOpus2ModelPath = settings.InputOpus2ModelPath;
                    InputTemplatePath = settings.InputTemplatePath;
                    OutputFilePath = settings.OutputFilePath;
                    Target = settings.Target;
                    InputTemplateDirectoryPath = settings.InputTemplateDirectoryPath;
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
                InputTemplateDirectoryPath = InputTemplateDirectoryPath
            };
            var json = JsonSerializer.Serialize(settings, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(GetSettingsFilePath(), json);
        }
        catch
        {
            // If saving fails, silently ignore
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

    [RelayCommand]
    private async Task GenerateDocumentAsync()
    {
        var da = new DocumentAssembler();

        var tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
        Directory.CreateDirectory(tempDir);
        try
        {
            var context = new DocumentAssembler.Context(InputInterfaceViewPath, InputDeploymentViewPath, Target, InputTemplateDirectoryPath, tempDir);
            await da.ProcessTemplate(context, InputTemplatePath, OutputFilePath);
        }
        finally
        {
            Directory.Delete(tempDir, true);
        }

    }
}
