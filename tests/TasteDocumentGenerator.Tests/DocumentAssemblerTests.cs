using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace TasteDocumentGenerator.Tests;

public class DocumentAssemblerTests
{
    private static string CreateMinimalDocx(string filePath, string? content = null)
    {
        using (var doc = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());

            if (content is not null)
            {
                var body = mainPart.Document.Body;
                var para = new Paragraph(new Run(new Text(content)));
                body!.Append(para);
            }
        }
        return filePath;
    }

    private static string CreateMinimalDocx(string filePath)
    {
        using (var doc = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());
        }
        return filePath;
    }

    [Fact]
    public void Context_ConstructorSetsAllProperties()
    {
        // Arrange
        var interfaceView = "iv.xml";
        var deploymentView = "dv.xml";
        var target = "ASW";
        var templateDir = "/templates";
        var tempDir = "/temp";
        var csvFiles = new[] { "memory.csv", "parameters.csv" };

        // Act
        var context = new DocumentAssembler.Context(
            interfaceView,
            deploymentView,
            target,
            templateDir,
            tempDir,
            null,
            null,
            csvFiles);

        // Assert
        Assert.Equal(interfaceView, context.InterfaceViewPath);
        Assert.Equal(deploymentView, context.DeploymentViewPath);
        Assert.Equal(target, context.Target);
        Assert.Equal(templateDir, context.TemplateDirectory);
        Assert.Equal(tempDir, context.TemporaryDirectory);
        Assert.Equal(csvFiles, context.SystemObjectCsvFiles);
    }

    [Fact]
    public void EnsureNumberingDefinitionsPart_CreatesPartWhenMissing()
    {
        // Arrange
        var tempFile = Path.GetTempFileName();
        try
        {
            CreateMinimalDocx(tempFile);
            using var doc = WordprocessingDocument.Open(tempFile, true);
            var assembler = new TestableDocumentAssembler();

            // Act
            var part = assembler.CallEnsureNumberingDefinitionsPart(doc);

            // Assert
            Assert.NotNull(part);
            Assert.NotNull(part.Numbering);
            Assert.Same(part, doc.MainDocumentPart!.NumberingDefinitionsPart);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void EnsureNumberingDefinitionsPart_ReturnsExistingPart()
    {
        // Arrange
        var tempFile = Path.GetTempFileName();
        try
        {
            CreateMinimalDocx(tempFile);
            using var doc = WordprocessingDocument.Open(tempFile, true);
            var existingPart = doc.MainDocumentPart!.AddNewPart<NumberingDefinitionsPart>();
            existingPart.Numbering = new Numbering();
            var assembler = new TestableDocumentAssembler();

            // Act
            var part = assembler.CallEnsureNumberingDefinitionsPart(doc);

            // Assert
            Assert.Same(existingPart, part);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void GetUsedAbstractIds_ReturnsEmptySetForEmptyNumbering()
    {
        // Arrange
        var tempFile = Path.GetTempFileName();
        try
        {
            CreateMinimalDocx(tempFile);
            using var doc = WordprocessingDocument.Open(tempFile, true);
            var part = doc.MainDocumentPart!.AddNewPart<NumberingDefinitionsPart>();
            part.Numbering = new Numbering();
            var assembler = new TestableDocumentAssembler();

            // Act
            var ids = assembler.CallGetUsedAbstractIds(part);

            // Assert
            Assert.Empty(ids);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void GetUsedAbstractIds_ReturnsAbstractNumberingIds()
    {
        // Arrange
        var tempFile = Path.GetTempFileName();
        try
        {
            CreateMinimalDocx(tempFile);
            using var doc = WordprocessingDocument.Open(tempFile, true);
            var part = doc.MainDocumentPart!.AddNewPart<NumberingDefinitionsPart>();
            part.Numbering = new Numbering();
            part.Numbering.Append(new AbstractNum() { AbstractNumberId = 0 });
            part.Numbering.Append(new AbstractNum() { AbstractNumberId = 5 });
            part.Numbering.Append(new AbstractNum() { AbstractNumberId = 10 });
            var assembler = new TestableDocumentAssembler();

            // Act
            var ids = assembler.CallGetUsedAbstractIds(part);

            // Assert
            Assert.Equal(3, ids.Count);
            Assert.Contains(0, ids);
            Assert.Contains(5, ids);
            Assert.Contains(10, ids);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void GetUsedNumberingIds_ReturnsEmptySetForEmptyNumbering()
    {
        // Arrange
        var tempFile = Path.GetTempFileName();
        try
        {
            CreateMinimalDocx(tempFile);
            using var doc = WordprocessingDocument.Open(tempFile, true);
            var part = doc.MainDocumentPart!.AddNewPart<NumberingDefinitionsPart>();
            part.Numbering = new Numbering();
            var assembler = new TestableDocumentAssembler();

            // Act
            var ids = assembler.CallGetUsedNumberingIds(part);

            // Assert
            Assert.Empty(ids);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void GetUsedNumberingIds_ReturnsNumberingInstanceIds()
    {
        // Arrange
        var tempFile = Path.GetTempFileName();
        try
        {
            CreateMinimalDocx(tempFile);
            using var doc = WordprocessingDocument.Open(tempFile, true);
            var part = doc.MainDocumentPart!.AddNewPart<NumberingDefinitionsPart>();
            part.Numbering = new Numbering();
            part.Numbering.Append(new NumberingInstance() { NumberID = 1 });
            part.Numbering.Append(new NumberingInstance() { NumberID = 3 });
            part.Numbering.Append(new NumberingInstance() { NumberID = 7 });
            var assembler = new TestableDocumentAssembler();

            // Act
            var ids = assembler.CallGetUsedNumberingIds(part);

            // Assert
            Assert.Equal(3, ids.Count);
            Assert.Contains(1, ids);
            Assert.Contains(3, ids);
            Assert.Contains(7, ids);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Theory]
    [InlineData("<TDG: template example />", "template example")]
    [InlineData("<TDG: command arg1 arg2 />", "command arg1 arg2")]
    [InlineData("<TDG:test/>", "test")]
    public void ExtractCommand_ParsesCorrectly(string input, string expected)
    {
        // Arrange
        var assembler = new TestableDocumentAssembler();

        // Act
        var result = assembler.CallExtractCommand(input, "<TDG:", "/>");

        // Assert
        Assert.Equal(expected, result);
    }

    [Fact]
    public void GetAllText_ConcatenatesTextFromParagraph()
    {
        // Arrange
        var tempFile = Path.GetTempFileName();
        try
        {
            CreateMinimalDocx(tempFile);
            using var doc = WordprocessingDocument.Open(tempFile, true);
            var body = doc.MainDocumentPart!.Document!.Body!;
            var para = new Paragraph();
            para.Append(new Run(new Text("Hello ")));
            para.Append(new Run(new Text("World")));
            body.Append(para);
            var assembler = new TestableDocumentAssembler();

            // Act
            var text = assembler.CallGetAllText(para);

            // Assert
            Assert.Equal("Hello World", text);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public async Task ProcessTemplate_WithNoHooks_GeneratesValidDocument()
    {
        // Arrange
        var tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
        Directory.CreateDirectory(tempDir);

        try
        {
            var templatePath = Path.Combine(tempDir, "template.docx");
            var outputPath = Path.Combine(tempDir, "output.docx");
            var ivPath = Path.Combine(tempDir, "iv.xml");
            var dvPath = Path.Combine(tempDir, "dv.xml");

            // Create minimal files
            CreateMinimalDocx(templatePath, "Simple template without hooks");
            File.WriteAllText(ivPath, "<InterfaceView/>");
            File.WriteAllText(dvPath, "<DeploymentView/>");

            var assembler = new DocumentAssembler();
            var context = new DocumentAssembler.Context(ivPath, dvPath, "ASW", tempDir, tempDir, null, null, Array.Empty<string>());

            // Act
            await assembler.ProcessTemplate(context, templatePath, outputPath);

            // Assert
            Assert.True(File.Exists(outputPath), "Output document should be created");

            // Verify the document can be opened and is valid
            using (var doc = WordprocessingDocument.Open(outputPath, false))
            {
                Assert.NotNull(doc.MainDocumentPart);
                Assert.NotNull(doc.MainDocumentPart.Document);
                Assert.NotNull(doc.MainDocumentPart.Document.Body);
            }
        }
        finally
        {
            Directory.Delete(tempDir, true);
        }
    }

    [Fact]
    public void InsertDocumentIntoParagraph_MergesDocuments()
    {
        // Arrange
        var tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
        Directory.CreateDirectory(tempDir);

        try
        {
            var targetPath = Path.Combine(tempDir, "target.docx");
            var sourcePath = Path.Combine(tempDir, "source.docx");

            CreateMinimalDocx(targetPath, "Target document");
            CreateMinimalDocx(sourcePath, "Source content to insert");

            var assembler = new DocumentAssembler();

            using (var targetDoc = WordprocessingDocument.Open(targetPath, true))
            {
                var body = targetDoc.MainDocumentPart!.Document!.Body!;
                var insertPara = body.Elements<Paragraph>().First();

                // Act
                assembler.InsertDocumentIntoParagraph(sourcePath, targetDoc, insertPara);
                targetDoc.Save();
            }

            // Assert
            using (var targetDoc = WordprocessingDocument.Open(targetPath, false))
            {
                var body = targetDoc.MainDocumentPart!.Document!.Body!;
                var paragraphs = body.Elements<Paragraph>().ToList();

                // Should have more than one paragraph now (original hook + inserted content)
                Assert.True(paragraphs.Count == 2, "Document should contain inserted content");
            }
        }
        finally
        {
            Directory.Delete(tempDir, true);
        }
    }

    [Fact]
    public void InsertDocumentIntoParagraph_WithNumberingAndStyles_MergesCorrectly()
    {
        // Arrange
        var tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
        Directory.CreateDirectory(tempDir);

        try
        {
            var targetPath = Path.Combine(tempDir, "target.docx");
            var sourcePath = Path.Combine(tempDir, "source.docx");

            // Create target document with numbering and styles
            using (var targetDoc = WordprocessingDocument.Create(targetPath, WordprocessingDocumentType.Document))
            {
                var mainPart = targetDoc.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());

                // Add numbering part with abstract numbering and instance
                var numberingPart = mainPart.AddNewPart<NumberingDefinitionsPart>();
                numberingPart.Numbering = new Numbering();
                var abstractNum = new AbstractNum() { AbstractNumberId = 0 };
                abstractNum.Append(new MultiLevelType() { Val = MultiLevelValues.HybridMultilevel });
                numberingPart.Numbering.Append(abstractNum);
                numberingPart.Numbering.Append(new NumberingInstance(
                    new AbstractNumId() { Val = 0 }
                )
                { NumberID = 1 });

                // Add styles part with a custom style
                var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
                stylesPart.Styles = new Styles();
                var customStyle = new Style()
                {
                    StyleId = "CustomStyle1",
                    Type = StyleValues.Paragraph
                };
                customStyle.Append(new StyleName() { Val = "Custom Style 1" });
                stylesPart.Styles.Append(customStyle);

                // Add target content with a placeholder paragraph
                var body = mainPart.Document.Body!;
                var placeholderPara = new Paragraph(new Run(new Text("PLACEHOLDER")));
                body.Append(placeholderPara);

                // Add a paragraph with numbering
                var numberedPara = new Paragraph();
                numberedPara.ParagraphProperties = new ParagraphProperties(
                    new NumberingProperties(
                        new NumberingLevelReference() { Val = 0 },
                        new NumberingId() { Val = 1 }
                    )
                );
                numberedPara.Append(new Run(new Text("Target numbered item")));
                body.Append(numberedPara);
            }

            // Create source document with different numbering and styles
            using (var sourceDoc = WordprocessingDocument.Create(sourcePath, WordprocessingDocumentType.Document))
            {
                var mainPart = sourceDoc.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());

                // Add numbering to source (will be remapped)
                var numberingPart = mainPart.AddNewPart<NumberingDefinitionsPart>();
                numberingPart.Numbering = new Numbering();
                var abstractNum = new AbstractNum() { AbstractNumberId = 0 };
                abstractNum.Append(new MultiLevelType() { Val = MultiLevelValues.HybridMultilevel });
                numberingPart.Numbering.Append(abstractNum);
                numberingPart.Numbering.Append(new NumberingInstance(
                    new AbstractNumId() { Val = 0 }
                )
                { NumberID = 1 });

                // Add styles to source
                var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
                stylesPart.Styles = new Styles();
                var sourceStyle = new Style()
                {
                    StyleId = "SourceStyle1",
                    Type = StyleValues.Paragraph
                };
                sourceStyle.Append(new StyleName() { Val = "Source Style 1" });
                stylesPart.Styles.Append(sourceStyle);

                // Add source content
                var body = mainPart.Document.Body!;

                var simplePara = new Paragraph(new Run(new Text("Simple paragraph from source")));
                body.Append(simplePara);

                var styledPara = new Paragraph();
                styledPara.ParagraphProperties = new ParagraphProperties(
                    new ParagraphStyleId() { Val = "SourceStyle1" }
                );
                styledPara.Append(new Run(new Text("Styled paragraph from source")));
                body.Append(styledPara);

                var numberedPara = new Paragraph();
                numberedPara.ParagraphProperties = new ParagraphProperties(
                    new NumberingProperties(
                        new NumberingLevelReference() { Val = 0 },
                        new NumberingId() { Val = 1 }
                    )
                );
                numberedPara.Append(new Run(new Text("Source numbered item")));
                body.Append(numberedPara);
            }

            var assembler = new DocumentAssembler();

            // Act - Insert source into target at placeholder
            using (var targetDoc = WordprocessingDocument.Open(targetPath, true))
            {
                var body = targetDoc.MainDocumentPart!.Document!.Body!;
                var placeholderPara = body.Elements<Paragraph>().First();

                assembler.InsertDocumentIntoParagraph(sourcePath, targetDoc, placeholderPara);
                targetDoc.Save();
            }

            // Assert
            using (var targetDoc = WordprocessingDocument.Open(targetPath, false))
            {
                var body = targetDoc.MainDocumentPart!.Document!.Body!;
                var paragraphs = body.Elements<Paragraph>().ToList();

                // Should have: empty placeholder + 3 source paragraphs + 1 target numbered = 5 total
                Assert.True(paragraphs.Count >= 4, $"Expected at least 4 paragraphs, got {paragraphs.Count}");

                // Check that numbering was merged
                var numberingPart = targetDoc.MainDocumentPart!.NumberingDefinitionsPart;
                Assert.NotNull(numberingPart);

                var abstractNums = numberingPart!.Numbering!.Elements<AbstractNum>().ToList();
                var numberingInstances = numberingPart.Numbering!.Elements<NumberingInstance>().ToList();

                // Should have 2 abstract numberings (target + source)
                Assert.Equal(2, abstractNums.Count);
                // Should have 2 numbering instances (target + source)
                Assert.Equal(2, numberingInstances.Count);

                // Check that styles were merged
                var stylesPart = targetDoc.MainDocumentPart!.StyleDefinitionsPart;
                Assert.NotNull(stylesPart);

                var styles = stylesPart!.Styles!.Elements<Style>().ToList();
                var styleIds = styles.Select(s => s.StyleId?.Value).Where(id => id != null).ToList();

                // Should contain both original and source styles
                Assert.Contains("CustomStyle1", styleIds);
                Assert.Contains("SourceStyle1", styleIds);

                // Verify text content was inserted
                var allText = string.Join(" ", paragraphs.Select(p =>
                    string.Concat(p.Descendants<Text>().Select(t => t.Text))));
                Assert.Contains("Simple paragraph from source", allText);
                Assert.Contains("Styled paragraph from source", allText);
                Assert.Contains("Source numbered item", allText);
            }
        }
        finally
        {
            Directory.Delete(tempDir, true);
        }
    }

    [Fact]
    public async Task ProcessTemplate_WithDocumentCommandAndImages_MergesImagesCorrectly()
    {
        // Arrange
        var tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
        Directory.CreateDirectory(tempDir);

        try
        {
            var templatePath = Path.Combine(tempDir, "template.docx");
            var outputPath = Path.Combine(tempDir, "output.docx");
            var sourceDocPath = Path.Combine(tempDir, "source-with-image.docx");
            var ivPath = Path.Combine(tempDir, "iv.xml");
            var dvPath = Path.Combine(tempDir, "dv.xml");

            // Create minimal files
            File.WriteAllText(ivPath, "<InterfaceView/>");
            File.WriteAllText(dvPath, "<DeploymentView/>");

            // Create a source document with an image
            using (var sourceDoc = WordprocessingDocument.Create(sourceDocPath, WordprocessingDocumentType.Document))
            {
                var mainPart = sourceDoc.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());
                var body = mainPart.Document.Body;

                // Add a simple paragraph with text
                body!.Append(new Paragraph(new Run(new Text("Document with image"))));

                // Add a minimal image (1x1 PNG)
                var imagePart = mainPart.AddImagePart(ImagePartType.Png);
                using (var stream = imagePart.GetStream())
                {
                    // Minimal 1x1 PNG
                    byte[] pngData = new byte[] {
                        0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
                        0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1F, 0x15, 0xC4,
                        0x89, 0x00, 0x00, 0x00, 0x0A, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9C, 0x63, 0x00, 0x01, 0x00, 0x00,
                        0x05, 0x00, 0x01, 0x0D, 0x0A, 0x2D, 0xB4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE,
                        0x42, 0x60, 0x82
                    };
                    stream.Write(pngData, 0, pngData.Length);
                }

                var imageRelId = mainPart.GetIdOfPart(imagePart);
                
                // Add a paragraph with the image
                var imageParagraph = new Paragraph();
                var run = new Run();
                var drawing = new DocumentFormat.OpenXml.Wordprocessing.Drawing(
                    new DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline(
                        new DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent() { Cx = 990000L, Cy = 990000L },
                        new DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties() { Id = 1U, Name = "Picture 1" },
                        new DocumentFormat.OpenXml.Drawing.Graphic(
                            new DocumentFormat.OpenXml.Drawing.GraphicData(
                                new DocumentFormat.OpenXml.Drawing.Pictures.Picture(
                                    new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureProperties(
                                        new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualDrawingProperties() { Id = 0U, Name = "Image.png" },
                                        new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureDrawingProperties()
                                    ),
                                    new DocumentFormat.OpenXml.Drawing.Pictures.BlipFill(
                                        new DocumentFormat.OpenXml.Drawing.Blip() { Embed = imageRelId },
                                        new DocumentFormat.OpenXml.Drawing.Stretch(
                                            new DocumentFormat.OpenXml.Drawing.FillRectangle()
                                        )
                                    ),
                                    new DocumentFormat.OpenXml.Drawing.Pictures.ShapeProperties()
                                )
                            ) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }
                        )
                    )
                );
                run.Append(drawing);
                imageParagraph.Append(run);
                body.Append(imageParagraph);

                // Add a caption paragraph
                body.Append(new Paragraph(new Run(new Text("Figure 1: Test Image"))));
            }

            // Create template with document command
            using (var templateDoc = WordprocessingDocument.Create(templatePath, WordprocessingDocumentType.Document))
            {
                var mainPart = templateDoc.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());
                var body = mainPart.Document.Body;
                body!.Append(new Paragraph(new Run(new Text("Template start"))));
                body.Append(new Paragraph(new Run(new Text("<TDG: document source-with-image.docx />"))));
                body.Append(new Paragraph(new Run(new Text("Template end"))));
            }

            var assembler = new DocumentAssembler();
            var context = new DocumentAssembler.Context(ivPath, dvPath, "ASW", tempDir, tempDir, null, null, Array.Empty<string>());

            // Act
            await assembler.ProcessTemplate(context, templatePath, outputPath);

            // Assert
            Assert.True(File.Exists(outputPath), "Output document should be created");

            using (var outputDoc = WordprocessingDocument.Open(outputPath, false))
            {
                var body = outputDoc.MainDocumentPart!.Document!.Body!;
                var paragraphs = body.Elements<Paragraph>().ToList();

                // Verify document content
                var allText = string.Join(" ", paragraphs.Select(p =>
                    string.Concat(p.Descendants<Text>().Select(t => t.Text))));
                Assert.Contains("Template start", allText);
                Assert.Contains("Document with image", allText);
                Assert.Contains("Figure 1: Test Image", allText);
                Assert.Contains("Template end", allText);

                // Verify image was copied
                var imageParts = outputDoc.MainDocumentPart.ImageParts.ToList();
                Assert.NotEmpty(imageParts);
                Assert.Single(imageParts);

                // Verify image references are valid
                var blips = body.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().ToList();
                Assert.NotEmpty(blips);
                foreach (var blip in blips)
                {
                    var embedId = blip.Embed?.Value;
                    Assert.NotNull(embedId);
                    var part = outputDoc.MainDocumentPart.GetPartById(embedId!);
                    Assert.NotNull(part);
                }
            }
        }
        finally
        {
            Directory.Delete(tempDir, true);
        }
    }

}

// Test helper class to expose private methods for testing
public class TestableDocumentAssembler : DocumentAssembler
{
    public NumberingDefinitionsPart CallEnsureNumberingDefinitionsPart(WordprocessingDocument target)
    {
        return EnsureNumberingDefinitionsPart(target);
    }

    public HashSet<int> CallGetUsedAbstractIds(NumberingDefinitionsPart part)
    {
        return GetUsedAbstractIds(part);
    }

    public HashSet<int> CallGetUsedNumberingIds(NumberingDefinitionsPart part)
    {
        return GetUsedNumberingIds(part);
    }

    public string CallExtractCommand(string text, string begin, string end)
    {
        return ExtractCommand(text, begin, end);
    }

    public string CallGetAllText(Paragraph paragraph)
    {
        return GetAllText(paragraph);
    }
}
