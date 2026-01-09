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
using Tmds.DBus.Protocol;

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

    public void InsertDocumentIntoParagraph(string path, WordprocessingDocument targetDocument, Paragraph paragraph)
    {
        paragraph.RemoveAllChildren<DocumentFormat.OpenXml.Wordprocessing.Run>();
        var parent = paragraph.Parent!;
        OpenXmlElement insertionPoint = paragraph;

        using (var sourceDocument = WordprocessingDocument.Open(path, false))
        {
            Debug.Print("Source document styles:");
            ListStyles(sourceDocument);
            var numberingIdMapping = MergeNumberingDefinitions(targetDocument, sourceDocument);
            var styleIdMapping = MergeDocumentStyles(targetDocument, sourceDocument, numberingIdMapping);
            var sourceBody = sourceDocument.MainDocumentPart?.Document.Body;
            if (sourceBody != null)
            {
                foreach (var element in sourceBody.Elements())
                {
                    var clonedElement = element.CloneNode(true);
                    UpdateParagraphNumbering(clonedElement, numberingIdMapping);
                    UpdateParagraphStyle(clonedElement, styleIdMapping);
                    GetStyle(clonedElement);
                    Debug.Print($"Cloning element {clonedElement.ToString()}");
                    parent.InsertAfter(clonedElement, insertionPoint);
                    insertionPoint = clonedElement;
                }
            }
        }
    }

    private void UpdateParagraphStyle(OpenXmlElement element, Dictionary<string, string> styleIdMapping)
    {
        if (element is Paragraph paragraph)
        {
            var styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            if (styleId != null && styleIdMapping.ContainsKey(styleId))
            {
                Debug.WriteLine($"Updating element style from {styleId} to {styleIdMapping[styleId]}");
                paragraph.ParagraphProperties!.ParagraphStyleId!.Val = styleIdMapping[styleId];
            }
        }

        if (element is Run run)
        {
            var styleId = run.RunProperties?.RunStyle?.Val?.Value;
            if (styleId != null && styleIdMapping.ContainsKey(styleId))
            {
                Debug.WriteLine($"Updating element style from {styleId} to {styleIdMapping[styleId]}");
                run.RunProperties!.RunStyle!.Val = styleIdMapping[styleId];
            }
        }

        foreach (var descendant in element.Descendants())
        {
            if (descendant is Paragraph p)
            {
                var styleId = p.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
                if (styleId != null && styleIdMapping.ContainsKey(styleId))
                {
                    Debug.WriteLine($"Updating element style from {styleId} to {styleIdMapping[styleId]}");
                    p.ParagraphProperties!.ParagraphStyleId!.Val = styleIdMapping[styleId];
                }
            }

            if (descendant is Run r)
            {
                var styleId = r.RunProperties?.RunStyle?.Val?.Value;
                if (styleId != null && styleIdMapping.ContainsKey(styleId))
                {
                    Debug.WriteLine($"Updating element style from {styleId} to {styleIdMapping[styleId]}");
                    r.RunProperties!.RunStyle!.Val = styleIdMapping[styleId];
                }
            }
        }
    }

    public async Task ProcessParagraphWithTemplate(Context context, WordprocessingDocument targetDocument, Paragraph paragraph, string[] command)
    {
        // First argument shall be a template
        var templatePath = command[1];
        if (command.Length != 2)
        {
            throw new Exception($"Incorrect template invocation: {string.Join(" ", command)}");
        }
        Debug.Print($"Processing template {command[1]}");
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

        InsertDocumentIntoParagraph(instancePath, targetDocument, paragraph);
    }

    public async Task ProcessParagraph(Context context, WordprocessingDocument targetDocument, Paragraph paragraph)
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
                    await ProcessParagraphWithTemplate(context, targetDocument, paragraph, command);
                    break;
                }
            default:
                {
                    break;
                }
        }
        ;
    }

    public void GetStyle(OpenXmlElement element)
    {
        var styleId = element.GetType().GetProperty("StyleId")?.GetValue(element) as StringValue;
        if (styleId?.Value != null)
        {
            Debug.Print($"Element style: {styleId.Value}");
        }

        if (element is Paragraph paragraph && paragraph.ParagraphProperties?.ParagraphStyleId?.Val != null)
        {
            Debug.Print($"Paragraph style: {paragraph.ParagraphProperties.ParagraphStyleId.Val.Value}");
        }

        if (element is Run run && run.RunProperties?.RunStyle?.Val != null)
        {
            Debug.Print($"Run style: {run.RunProperties.RunStyle.Val.Value}");
        }

        foreach (var descendant in element.Descendants())
        {
            if (descendant is Paragraph p && p.ParagraphProperties?.ParagraphStyleId?.Val != null)
            {
                Debug.Print($"Descendant paragraph style: {p.ParagraphProperties.ParagraphStyleId.Val.Value}");
            }
            if (descendant is Run r && r.RunProperties?.RunStyle?.Val != null)
            {
                Debug.Print($"Descendant run style: {r.RunProperties.RunStyle.Val.Value}");
            }
        }
    }

    public Dictionary<string, string> MergeDocumentStyles(WordprocessingDocument target, WordprocessingDocument source, Dictionary<int, int> numberingIdMapping)
    {
        var mapping = new Dictionary<string, string>();
        var targetStylesPart = target.MainDocumentPart?.StyleDefinitionsPart;
        var targetStylesWithEffectsPart = target.MainDocumentPart?.StylesWithEffectsPart;
        var sourceStylesPart = source.MainDocumentPart?.StyleDefinitionsPart;
        var sourceStylesWithEffectsPart = source.MainDocumentPart?.StylesWithEffectsPart;

        // Collect all existing style IDs in target document
        var existingStyleIds = new HashSet<string>();

        if (targetStylesPart?.Styles != null)
        {
            foreach (var style in targetStylesPart.Styles.Elements<Style>())
            {
                if (style.StyleId?.Value != null)
                {
                    existingStyleIds.Add(style.StyleId.Value);
                }
            }
        }

        if (targetStylesWithEffectsPart?.Styles != null)
        {
            foreach (var style in targetStylesWithEffectsPart.Styles.Elements<Style>())
            {
                if (style.StyleId?.Value != null)
                {
                    existingStyleIds.Add(style.StyleId.Value);
                }
            }
        }

        // Ensure target has StyleDefinitionsPart
        if (targetStylesPart == null)
        {
            targetStylesPart = target.MainDocumentPart!.AddNewPart<StyleDefinitionsPart>();
            targetStylesPart.Styles = new Styles();
        }
        if (targetStylesWithEffectsPart == null)
        {
            targetStylesWithEffectsPart = target.MainDocumentPart!.AddNewPart<StylesWithEffectsPart>();
            targetStylesWithEffectsPart.Styles = new Styles();
        }

        // Copy styles from source Styles that don't exist in target
        if (sourceStylesPart?.Styles != null)
        {
            foreach (var style in sourceStylesPart.Styles.Elements<Style>())
            {
                if (style.StyleId?.Value != null)
                {
                    var clonedStyle = (Style)style.CloneNode(true);
                    if (existingStyleIds.Contains(style.StyleId.Value))
                    {
                        var oldId = style.StyleId.Value;
                        var newId = oldId + "src";
                        clonedStyle?.StyleId?.Value = newId;
                        Debug.WriteLine($"Mapping style {oldId} to {newId}");
                        mapping[oldId] = newId;
                    }
                    var numberingId = style.StyleParagraphProperties?.NumberingProperties?.NumberingId?.Val;
                    if (numberingId != null && numberingIdMapping.ContainsKey(numberingId))
                    {
                        Debug.WriteLine($"Updating style {clonedStyle!.StyleId!.Value} numbering from {numberingId} to {numberingIdMapping[numberingId]}");
                        clonedStyle!.StyleParagraphProperties!.NumberingProperties!.NumberingId!.Val = numberingIdMapping[numberingId];
                    }
                    targetStylesPart?.Styles?.Append(clonedStyle);
                    existingStyleIds.Add(clonedStyle?.StyleId?.Value ?? "");
                }
            }
        }

        // Copy styles from source StylesWithEffects that don't exist in target
        if (sourceStylesWithEffectsPart?.Styles != null)
        {
            foreach (var style in sourceStylesWithEffectsPart.Styles.Elements<Style>())
            {
                if (style.StyleId?.Value != null)
                {
                    var clonedStyle = (Style)style.CloneNode(true);
                    if (existingStyleIds.Contains(style.StyleId.Value))
                    {
                        var oldId = style.StyleId.Value;
                        var newId = oldId + "src";
                        clonedStyle?.StyleId?.Value = newId;
                        Debug.WriteLine($"Mapping style {oldId} to {newId}");
                        mapping[oldId] = newId;
                    }
                    var numberingId = style.StyleParagraphProperties?.NumberingProperties?.NumberingId?.Val;
                    if (numberingId != null && numberingIdMapping.ContainsKey(numberingId))
                    {

                        Debug.WriteLine($"Updating style {clonedStyle!.StyleId!.Value} numbering from {numberingId} to {numberingIdMapping[numberingId]}");
                        clonedStyle!.StyleParagraphProperties!.NumberingProperties!.NumberingId!.Val = numberingIdMapping[numberingId];
                    }
                    targetStylesWithEffectsPart?.Styles?.Append(clonedStyle);
                    existingStyleIds.Add(style.StyleId.Value);
                }
            }
        }
        return mapping;
    }

    public void ListStyles(WordprocessingDocument document)
    {
        var stylesPart = document.MainDocumentPart?.StyleDefinitionsPart;
        var stylesWithEffectsPart = document.MainDocumentPart?.StylesWithEffectsPart;
        if (stylesPart?.Styles == null && stylesWithEffectsPart?.Styles == null)
        {
            Debug.Print("No styles found in document");
            return;
        }

        if (stylesPart?.Styles != null)
        {
            foreach (var style in stylesPart.Styles.Elements<Style>())
            {
                var styleId = style.StyleId?.Value ?? "N/A";
                var styleName = style.StyleName?.Val?.Value ?? "N/A";
                Debug.Print($"Style ID: {styleId}, Name: {styleName}");
            }
        }
        if (stylesWithEffectsPart?.Styles != null)
        {
            foreach (var style in stylesWithEffectsPart.Styles.Elements<Style>())
            {
                var styleId = style.StyleId?.Value ?? "N/A";
                var styleName = style.StyleName?.Val?.Value ?? "N/A";
                Debug.Print($"Style ID: {styleId}, Name: {styleName}");
            }
        }
    }




    private Dictionary<int, int> MergeNumberingDefinitions(WordprocessingDocument target, WordprocessingDocument source)
    {
        var mapping = new Dictionary<int, int>();
        var abstractNumberingMapping = new Dictionary<int, int>();

        var sourceNumbering = source.MainDocumentPart?.NumberingDefinitionsPart;
        if (sourceNumbering?.Numbering == null)
            return mapping;

        var targetNumbering = target.MainDocumentPart?.NumberingDefinitionsPart;
        if (targetNumbering == null)
        {
            targetNumbering = target.MainDocumentPart!.AddNewPart<NumberingDefinitionsPart>();
            targetNumbering.Numbering = new Numbering();
        }

        var usedAbstractIds = new HashSet<int>();
        foreach (var abstractNumbering in targetNumbering.Numbering.Elements<AbstractNum>())
        {
            if (abstractNumbering.AbstractNumberId?.Value != null)
                usedAbstractIds.Add(abstractNumbering.AbstractNumberId.Value);
        }

        var usedIds = new HashSet<int>();
        foreach (var numbering in targetNumbering.Numbering.Elements<NumberingInstance>())
        {
            if (numbering.NumberID?.Value != null)
                usedIds.Add(numbering.NumberID.Value);
        }

        int nextAbstractId = usedAbstractIds.Any() ? usedAbstractIds.Max() + 1 : 0;
        int nextNumId = usedIds.Any() ? usedIds.Max() + 1 : 1;


        foreach (var sourceAbstractNumberingInstance in sourceNumbering.Numbering.Elements<AbstractNum>())
        {
            if (sourceAbstractNumberingInstance.AbstractNumberId?.Value != null)
            {
                int oldAbstractId = sourceAbstractNumberingInstance.AbstractNumberId.Value;
                int newAbstractId = nextAbstractId++;

                var cloned = (AbstractNum)sourceAbstractNumberingInstance.CloneNode(true);
                cloned.AbstractNumberId = newAbstractId;
                targetNumbering.Numbering.Append(cloned);

                abstractNumberingMapping[oldAbstractId] = newAbstractId;
            }
        }

        foreach (var sourceNumberingInstance in sourceNumbering.Numbering.Elements<NumberingInstance>())
        {
            if (sourceNumberingInstance.NumberID?.Value != null)
            {
                int oldId = sourceNumberingInstance.NumberID.Value;
                int newNumId = nextNumId++;

                var cloned = (NumberingInstance)sourceNumberingInstance.CloneNode(true);
                cloned.NumberID = newNumId;

                var abstractRef = cloned.GetFirstChild<AbstractNumId>();
                if (abstractRef?.Val?.Value != null && abstractNumberingMapping.ContainsKey(abstractRef.Val.Value))
                {
                    abstractRef.Val = abstractNumberingMapping[abstractRef.Val.Value];
                }

                targetNumbering.Numbering.Append(cloned);
                mapping[oldId] = newNumId;
            }
        }

        return mapping;
    }

    private void UpdateParagraphNumbering(OpenXmlElement element, Dictionary<int, int> mapping)
    {
        foreach (var paragraph in element.Descendants<Paragraph>())
        {
            var property = paragraph.ParagraphProperties?.NumberingProperties;
            if (property?.NumberingId?.Val?.Value != null)
            {
                int oldId = property.NumberingId.Val.Value;
                if (mapping.ContainsKey(oldId))
                {
                    property.NumberingId.Val = mapping[oldId];
                }
            }
        }
    }


    public async Task ProcessTemplate(Context context, string inputTemplatePath, string outputDocumentPath)
    {
        File.Copy(inputTemplatePath, outputDocumentPath, true);

        using (var document = WordprocessingDocument.Open(outputDocumentPath, true))
        {
            Debug.Print("Target document styles:");
            ListStyles(document);
            var hooks = FindHooks(document, "template");

            foreach (var paragraph in hooks)
            {
                await ProcessParagraph(context, document, paragraph);
            }

            document.Save();
        }

        return;
    }
}