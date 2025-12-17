namespace TasteDocumentGenerator;

using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

class DocumentAssembler
{

    public const string BEGIN = "<TDG:";
    public const string END = "/>";

    public class Context
    {
        public string InterfaceViewPath;
        public string DeploymentViewPath;

        public Context(string InterfaceViewPath, string DeploymentViewPath)
        {
            this.InterfaceViewPath = InterfaceViewPath;
            this.DeploymentViewPath = DeploymentViewPath;
        }
    }

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
            var text = string.Concat(paragraph.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>().Select(t => t.Text)).Trim();
            if (text.StartsWith(BEGIN) && text.EndsWith(END))
            {
                var content = text.Substring(BEGIN.Length, text.Length - (BEGIN.Length + END.Length)).Trim();
                if (content.StartsWith(prefix))
                {
                    hooks.Add(paragraph);
                }
            }
        }
        
        return hooks;
    }

    public async Task ProcessTemplate(Context context, string inputTemplatePath, string outputDocumentPath)
    {
        File.Copy(inputTemplatePath, outputDocumentPath, true);

        using (var document = WordprocessingDocument.Open(outputDocumentPath, true))
        {
            var hooks = FindHooks(document, "template");
            
            foreach (var paragraph in hooks)
            {
                paragraph.RemoveAllChildren<DocumentFormat.OpenXml.Wordprocessing.Run>();
                paragraph.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Run(
                    new DocumentFormat.OpenXml.Wordprocessing.Text("FOUND!")));
            }
            
            document.Save();
        }

        return;
    }
}