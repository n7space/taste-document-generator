namespace TasteDocumentGenerator;

using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

class DocumentAssembler
{

    public const string BEGIN = "<TDG:";
    public const string END = "/>";

    public class Context
    {
        public string Target;
        public string TemplateDirectory;
        public string InterfaceViewPath;
        public string DeploymentViewPath;
        public string TemporaryDirectory;

        public Context(string InterfaceViewPath, string DeploymentViewPath, string Target, string TemplateDirectory, string TemporaryDirectory)
        {
            this.InterfaceViewPath = InterfaceViewPath;
            this.DeploymentViewPath = DeploymentViewPath;
            this.TemporaryDirectory = TemporaryDirectory;
            this.TemplateDirectory = TemplateDirectory;
            this.Target = Target;
        }
    }

    private string GetAllText(Paragraph paragraph) =>
        string.Concat(paragraph.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>().Select(t => t.Text)).Trim();


    private string ExtractCommand(string text) => text.Substring(BEGIN.Length, text.Length - (BEGIN.Length + END.Length)).Trim();

    private List<DocumentFormat.OpenXml.Wordprocessing.Paragraph> FindHooks(WordprocessingDocument document, string prefix)
    {
        var hooks = new List<DocumentFormat.OpenXml.Wordprocessing.Paragraph>();
        var body = document.MainDocumentPart?.Document.Body;

        if (body is null)
        {
            return [];
        }
        foreach (var paragraph in body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Paragraph>())
        {
            var text = GetAllText(paragraph);
            if (text.StartsWith(BEGIN) && text.EndsWith(END))
            {
                var content = ExtractCommand(text);
                if (content.StartsWith(prefix))
                {
                    hooks.Add(paragraph);
                }
            }
        }

        return hooks;
    }

    public void InsertDocumentIntoParagraph(string path, Paragraph paragraph)
    {
        paragraph.RemoveAllChildren<DocumentFormat.OpenXml.Wordprocessing.Run>();
        var parent = paragraph.Parent!;
        OpenXmlElement insertionPoint = paragraph;

        using (var sourceDocument = WordprocessingDocument.Open(path, false))
        {
            var sourceBody = sourceDocument.MainDocumentPart?.Document.Body;
            if (sourceBody != null)
            {
                foreach (var element in sourceBody.Elements())
                {
                    var clonedElement = element.CloneNode(true);
                    parent.InsertAfter(clonedElement, insertionPoint);
                    insertionPoint = clonedElement;
                }
            }
        }
    }

    public async Task ProcessParagraphWithTemplate(Context context, Paragraph paragraph, string[] command)
    {
        // First argument shall be a template
        var templatePath = command[1];
        if (command.Length != 2)
        {
            throw new Exception($"Incorrect template invocation: {string.Join(" ", command)}");
        }
        var processInfo = new ProcessStartInfo
        {
            FileName = "template-processor",
            Arguments = $" --verbosity info --iv {context.InterfaceViewPath} --dv {context.DeploymentViewPath} -o {context.TemporaryDirectory} -t {templatePath} -p md2docx --value TARGET=ASW",
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = false
        };

        using var process = Process.Start(processInfo);
        if (process is null)
        {
            throw new Exception("Could not start template-procesoor");
        }
        var outputTask = process.StandardOutput.ReadToEndAsync();
        var errorTask = process.StandardError.ReadToEndAsync();

        await process.WaitForExitAsync();

        var standardOutput = await outputTask;
        var standardError = await errorTask;
        var processOutput = standardOutput + standardError;


        var baseName = Path.GetFileNameWithoutExtension(templatePath);
        var instancePath = Path.Join(context.TemporaryDirectory, $"{baseName}.docx");

        if (!Path.Exists(instancePath))
        {
            throw new Exception($"File {instancePath} does not exist, did template instantiation fail? Template instantiation process output: {processOutput}");
        }

        InsertDocumentIntoParagraph(instancePath, paragraph);
    }

    public async Task ProcessParagraph(Context context, Paragraph paragraph)
    {
        var text = GetAllText(paragraph);
        var command = ExtractCommand(text).Split(" ");
        if (command.Length == 0)
        {
            return;
        }
        var commandName = command[0];
        switch (commandName)
        {
            case "template":
                {
                    await ProcessParagraphWithTemplate(context, paragraph, command);
                    break;
                }
            default:
                {
                    break;
                }
        }
        ;
    }

    public async Task ProcessTemplate(Context context, string inputTemplatePath, string outputDocumentPath)
    {
        File.Copy(inputTemplatePath, outputDocumentPath, true);

        using (var document = WordprocessingDocument.Open(outputDocumentPath, true))
        {
            var hooks = FindHooks(document, "template");

            foreach (var paragraph in hooks)
            {
                await ProcessParagraph(context, paragraph);
            }

            document.Save();
        }

        return;
    }
}