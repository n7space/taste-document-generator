using System;
using System.Threading.Tasks;
using System.Windows.Input;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Platform.Storage;
using Avalonia.Controls.ApplicationLifetimes;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;

namespace TasteDocumentGenerator;

public partial class MainWindowViewModel : ObservableObject
{
    [ObservableProperty]
    private string _inputInterfaceViewPath = "interfaceview.xml";

    [ObservableProperty]
    private string _inputDeploymentViewPath = "deploymentview.dv.xml";

    [ObservableProperty]
    private string _inputOpus2ModelPath = "opus2_model.xml";

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

        var file = await storageProvider.SaveFilePickerAsync(new FilePickerSaveOptions
        {
            Title = "Select input template path",
            DefaultExtension = "docx",
            SuggestedFileName = InputTemplatePath,
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
            InputTemplatePath = file.Path.LocalPath;
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

        await Task.CompletedTask;
    }
}
