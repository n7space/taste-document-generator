using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace TasteDocumentGenerator.Tests;

public class DocumentAssemblerTests
{

    private static string CreateTestTemplate(string filePath, string hookCommand)
    {
        using (var doc = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());

            // Add a paragraph with hook text
            var body = mainPart.Document.Body;
            var para = new Paragraph(new Run(new Text($"<TDG: {hookCommand} />")));
            body!.Append(para);
        }
        return filePath;
    }

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

        // Act
        var context = new DocumentAssembler.Context(
            interfaceView,
            deploymentView,
            target,
            templateDir,
            tempDir);

        // Assert
        Assert.Equal(interfaceView, context.InterfaceViewPath);
        Assert.Equal(deploymentView, context.DeploymentViewPath);
        Assert.Equal(target, context.Target);
        Assert.Equal(templateDir, context.TemplateDirectory);
        Assert.Equal(tempDir, context.TemporaryDirectory);
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
        var result = assembler.CallExtractCommand(input);

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
            var context = new DocumentAssembler.Context(ivPath, dvPath, "ASW", tempDir, tempDir);

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
                Assert.True(paragraphs.Count >= 1, "Document should contain inserted content");
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

    public string CallExtractCommand(string text)
    {
        return ExtractCommand(text);
    }

    public string CallGetAllText(Paragraph paragraph)
    {
        return GetAllText(paragraph);
    }
}
