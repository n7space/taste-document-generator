namespace TasteDocumentGenerator;

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

public interface IDocumentAssembler
{
    Task ProcessTemplate(DocumentAssembler.Context context, string inputTemplatePath, string outputDocumentPath);
}

public class DocumentAssembler : IDocumentAssembler
{
    public readonly string MARKER_OPEN = "<";
    public readonly string MARKER_CLOSE = "/>";

    public class Context
    {
        public readonly string? Target;
        public readonly string TemplateDirectory;
        public readonly string InterfaceViewPath;
        public readonly string DeploymentViewPath;
        public readonly string TemporaryDirectory;
        public readonly string? TemplateProcessor;
        public readonly string Tag;
        public IReadOnlyList<string> SystemObjectCsvFiles { get; }

        public Context(string InterfaceViewPath, string DeploymentViewPath, string? Target, string TemplateDirectory, string TemporaryDirectory, string? TemplateProcessor, string? Tag = null, IEnumerable<string>? systemObjectCsvFiles = null)
        {
            this.InterfaceViewPath = InterfaceViewPath;
            this.DeploymentViewPath = DeploymentViewPath;
            this.TemporaryDirectory = TemporaryDirectory;
            this.TemplateDirectory = TemplateDirectory;
            this.Target = Target;
            this.TemplateProcessor = TemplateProcessor;
            this.Tag = Tag ?? "TDG:";
            SystemObjectCsvFiles = systemObjectCsvFiles?.Where(path => !string.IsNullOrWhiteSpace(path)).Select(path => path.Trim()).ToArray() ?? Array.Empty<string>();
        }
    }

    protected string GetAllText(Paragraph paragraph) =>
        string.Concat(paragraph.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>().Select(t => t.Text)).Trim();


    protected string ExtractCommand(string text, string begin, string end) => text.Substring(begin.Length, text.Length - (begin.Length + end.Length)).Trim();

    private List<DocumentFormat.OpenXml.Wordprocessing.Paragraph> FindHooks(WordprocessingDocument document, string prefix, string begin, string end)
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
            if (text.StartsWith(begin) && text.EndsWith(end))
            {
                var content = ExtractCommand(text, begin, end);
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
            var numberingIdMapping = MergeNumberingDefinitions(targetDocument, sourceDocument);
            var styleIdMapping = MergeDocumentStyles(targetDocument, sourceDocument, numberingIdMapping);
            var imagePartMapping = MergeImageParts(targetDocument, sourceDocument);
            var sourceBody = sourceDocument.MainDocumentPart?.Document.Body;
            if (sourceBody is null)
            {
                Debug.WriteLine($"Body is null in target document, insertion skipped");
                return;
            }
            foreach (var element in sourceBody.Elements())
            {
                var clonedElement = element.CloneNode(true);
                UpdateParagraphNumbering(clonedElement, numberingIdMapping);
                UpdateParagraphStyle(clonedElement, styleIdMapping);
                UpdateImageReferences(clonedElement, imagePartMapping);
                parent.InsertAfter(clonedElement, insertionPoint);
                insertionPoint = clonedElement;
            }
        }
    }

    private void UpdateParagraphStyle(OpenXmlElement element, Dictionary<string, string> styleIdMapping)
    {
        void UpdateStyleOnElement(OpenXmlElement nestedElement)
        {
            if (nestedElement is Paragraph p)
            {
                var sid = p.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
                if (sid != null && styleIdMapping.ContainsKey(sid))
                {
                    Debug.WriteLine($"Updating element style from {sid} to {styleIdMapping[sid]}");
                    p.ParagraphProperties!.ParagraphStyleId!.Val = styleIdMapping[sid];
                }
            }

            if (nestedElement is Run r)
            {
                var sid = r.RunProperties?.RunStyle?.Val?.Value;
                if (sid != null && styleIdMapping.ContainsKey(sid))
                {
                    Debug.WriteLine($"Updating element style from {sid} to {styleIdMapping[sid]}");
                    r.RunProperties!.RunStyle!.Val = styleIdMapping[sid];
                }
            }
        }

        UpdateStyleOnElement(element);
        foreach (var descendant in element.Descendants())
        {
            UpdateStyleOnElement(descendant);
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
            FileName = context.TemplateProcessor ?? "template-processor",
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = false
        };

        processInfo.ArgumentList.Add("--verbosity");
        processInfo.ArgumentList.Add("info");
        processInfo.ArgumentList.Add("--iv");
        processInfo.ArgumentList.Add(context.InterfaceViewPath);
        processInfo.ArgumentList.Add("--dv");
        processInfo.ArgumentList.Add(context.DeploymentViewPath);
        processInfo.ArgumentList.Add("-o");
        processInfo.ArgumentList.Add(context.TemporaryDirectory);
        processInfo.ArgumentList.Add("-t");
        processInfo.ArgumentList.Add(Path.Join(context.TemplateDirectory, templatePath));
        processInfo.ArgumentList.Add("-p");
        processInfo.ArgumentList.Add("md2docx");
        if (!string.IsNullOrWhiteSpace(context.Target))
        {
            processInfo.ArgumentList.Add("--value");
            processInfo.ArgumentList.Add($"TARGET={context.Target}");
        }

        foreach (var csvPath in context.SystemObjectCsvFiles)
        {
            processInfo.ArgumentList.Add("-s");
            processInfo.ArgumentList.Add(csvPath);
        }

        var invocationArguments = string.Join(" ", processInfo.ArgumentList);
        Debug.WriteLine($"Calling {processInfo.FileName} with arguments {invocationArguments}");

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

    public void ProcessParagraphWithDocument(Context context, WordprocessingDocument targetDocument, Paragraph paragraph, string[] command)
    {
        if (command.Length != 2)
        {
            throw new Exception($"Incorrect document invocation: {string.Join(" ", command)}");
        }
        var documentPath = command[1];
        Debug.WriteLine($"Processing document {documentPath}");
        var fullPath = Path.IsPathRooted(documentPath)
            ? documentPath
            : Path.Join(context.TemplateDirectory, documentPath);

        if (!File.Exists(fullPath))
        {
            throw new Exception($"Document file {fullPath} does not exist");
        }

        InsertDocumentIntoParagraph(fullPath, targetDocument, paragraph);
    }

    public async Task ProcessParagraph(Context context, WordprocessingDocument targetDocument, Paragraph paragraph)
    {
        var begin = MARKER_OPEN + context.Tag;
        var end = MARKER_CLOSE;
        var text = GetAllText(paragraph);
        var command = ExtractCommand(text, begin, end).Split(" ");
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
            case "document":
                {
                    ProcessParagraphWithDocument(context, targetDocument, paragraph, command);
                    break;
                }
            default:
                {
                    break;
                }
        }
    }

    private HashSet<string> GetExistingStyleIds(StyleDefinitionsPart targetStylesPart, StylesWithEffectsPart targetStylesWithEffectsPart)
    {
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
        return existingStyleIds;
    }

    private Style CloneAndRemapStyle(Style style, HashSet<string> existingStyleIds, Dictionary<int, int> numberingIdMapping, Dictionary<string, string> styleIdMapping)
    {
        var clonedStyle = (Style)style.CloneNode(true);
        if (existingStyleIds.Contains(style!.StyleId!.Value!))
        {
            var oldId = style.StyleId.Value!;
            var newId = oldId + "remapped";
            clonedStyle?.StyleId?.Value = newId;
            Debug.WriteLine($"Mapping style {oldId} to {newId}");
            styleIdMapping[oldId] = newId;
        }
        var numberingId = style.StyleParagraphProperties?.NumberingProperties?.NumberingId?.Val;
        if (numberingId != null && numberingIdMapping.ContainsKey(numberingId))
        {
            Debug.WriteLine($"Updating style {clonedStyle!.StyleId!.Value} numbering from {numberingId} to {numberingIdMapping[numberingId]}");
            clonedStyle!.StyleParagraphProperties!.NumberingProperties!.NumberingId!.Val = numberingIdMapping[numberingId];
        }
        existingStyleIds.Add(clonedStyle?.StyleId?.Value ?? "");
        return clonedStyle!;
    }

    public Dictionary<string, string> MergeDocumentStyles(WordprocessingDocument target, WordprocessingDocument source, Dictionary<int, int> numberingIdMapping)
    {
        var mapping = new Dictionary<string, string>();
        var targetStylesPart = target.MainDocumentPart?.StyleDefinitionsPart;
        var targetStylesWithEffectsPart = target.MainDocumentPart?.StylesWithEffectsPart;
        var sourceStylesPart = source.MainDocumentPart?.StyleDefinitionsPart;
        var sourceStylesWithEffectsPart = source.MainDocumentPart?.StylesWithEffectsPart;

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
        var existingStyleIds = GetExistingStyleIds(targetStylesPart, targetStylesWithEffectsPart);

        if (sourceStylesPart?.Styles != null)
        {
            foreach (var style in sourceStylesPart.Styles.Elements<Style>())
            {
                if (style.StyleId?.Value != null)
                {
                    var clonedStyle = CloneAndRemapStyle(style, existingStyleIds, numberingIdMapping, mapping);
                    targetStylesPart?.Styles?.Append(clonedStyle);
                }
            }
        }

        if (sourceStylesWithEffectsPart?.Styles != null)
        {
            foreach (var style in sourceStylesWithEffectsPart.Styles.Elements<Style>())
            {
                if (style.StyleId?.Value != null)
                {
                    var clonedStyle = CloneAndRemapStyle(style, existingStyleIds, numberingIdMapping, mapping);
                    targetStylesWithEffectsPart?.Styles?.Append(clonedStyle);
                }
            }
        }
        return mapping;
    }

    private Dictionary<int, int> MergeNumberingDefinitions(WordprocessingDocument target, WordprocessingDocument source)
    {
        var mapping = new Dictionary<int, int>();
        var abstractNumberingMapping = new Dictionary<int, int>();

        var sourceNumbering = source.MainDocumentPart?.NumberingDefinitionsPart;
        if (sourceNumbering?.Numbering == null)
            return mapping;

        var targetNumbering = EnsureNumberingDefinitionsPart(target);

        var usedAbstractIds = GetUsedAbstractIds(targetNumbering);
        var usedIds = GetUsedNumberingIds(targetNumbering);

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

    protected NumberingDefinitionsPart EnsureNumberingDefinitionsPart(WordprocessingDocument target)
    {
        var part = target.MainDocumentPart?.NumberingDefinitionsPart;
        if (part == null)
        {
            part = target.MainDocumentPart!.AddNewPart<NumberingDefinitionsPart>();
            part.Numbering = new Numbering();
        }
        else if (part.Numbering == null)
        {
            part.Numbering = new Numbering();
        }
        return part!;
    }

    protected HashSet<int> GetUsedAbstractIds(NumberingDefinitionsPart numberingPart)
    {
        var used = new HashSet<int>();
        if (numberingPart?.Numbering == null)
            return used;

        foreach (var abstractNum in numberingPart.Numbering.Elements<AbstractNum>())
        {
            if (abstractNum.AbstractNumberId?.Value != null)
                used.Add(abstractNum.AbstractNumberId.Value);
        }
        return used;
    }

    protected HashSet<int> GetUsedNumberingIds(NumberingDefinitionsPart numberingPart)
    {
        var used = new HashSet<int>();
        if (numberingPart?.Numbering == null)
            return used;

        foreach (var numInstance in numberingPart.Numbering.Elements<NumberingInstance>())
        {
            if (numInstance.NumberID?.Value != null)
                used.Add(numInstance.NumberID.Value);
        }
        return used;
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

    private Dictionary<string, string> MergeImageParts(WordprocessingDocument target, WordprocessingDocument source)
    {
        var mapping = new Dictionary<string, string>();
        var targetMainPart = target.MainDocumentPart;
        var sourceMainPart = source.MainDocumentPart;

        if (sourceMainPart == null || targetMainPart == null)
        {
            return mapping;
        }

        foreach (var imagePart in sourceMainPart.ImageParts)
        {
            var sourceRelId = sourceMainPart.GetIdOfPart(imagePart);
            var newImagePart = CopyImagePart(targetMainPart, imagePart);
            var targetRelId = targetMainPart.GetIdOfPart(newImagePart);
            mapping[sourceRelId] = targetRelId;
        }

        return mapping;
    }

    private ImagePart CopyImagePart(MainDocumentPart targetMainPart, ImagePart sourceImagePart)
    {
        var contentType = sourceImagePart.ContentType;
        var newImagePart = CreateImagePart(targetMainPart, contentType);

        using (var sourceStream = sourceImagePart.GetStream())
        using (var targetStream = newImagePart.GetStream())
        {
            sourceStream.CopyTo(targetStream);
        }

        return newImagePart;
    }

    private static PartTypeInfo MimeToImageType(string contentType)
    {
        if (string.IsNullOrWhiteSpace(contentType))
            return ImagePartType.Jpeg;

        return contentType.ToLowerInvariant() switch
        {
            "image/bmp" => ImagePartType.Bmp,
            "image/gif" => ImagePartType.Gif,
            "image/png" => ImagePartType.Png,
            "image/tiff" => ImagePartType.Tiff,
            "image/x-icon" => ImagePartType.Icon,
            "image/x-pcx" => ImagePartType.Pcx,
            "image/jpeg" => ImagePartType.Jpeg,
            "image/x-emf" => ImagePartType.Emf,
            "image/x-wmf" => ImagePartType.Wmf,
            _ => ImagePartType.Jpeg
        };
    }

    private ImagePart CreateImagePart(MainDocumentPart mainPart, string contentType)
    {
        var imageType = MimeToImageType(contentType);
        return mainPart.AddImagePart(imageType);
    }

    private void UpdateImageReferences(OpenXmlElement element, Dictionary<string, string> imagePartMapping)
    {
        foreach (var blip in element.Descendants<DocumentFormat.OpenXml.Drawing.Blip>())
        {
            var embedId = blip.Embed?.Value;
            if (embedId != null && imagePartMapping.ContainsKey(embedId))
            {
                blip.Embed = imagePartMapping[embedId];
            }
        }
    }


    public async Task ProcessTemplate(Context context, string inputTemplatePath, string outputDocumentPath)
    {
        var begin = MARKER_OPEN + context.Tag;
        var end = MARKER_CLOSE;
        File.Copy(inputTemplatePath, outputDocumentPath, true);
        using (var document = WordprocessingDocument.Open(outputDocumentPath, true))
        {
            var templateHooks = FindHooks(document, "template", begin, end);
            var documentHooks = FindHooks(document, "document", begin, end);

            foreach (var paragraph in templateHooks)
            {
                await ProcessParagraph(context, document, paragraph);
            }

            foreach (var paragraph in documentHooks)
            {
                await ProcessParagraph(context, document, paragraph);
            }

            document.Save();
        }
        return;
    }
}