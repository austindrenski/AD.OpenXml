using System.IO;
using System.Linq;
using AD.IO;
using AD.OpenXml.Documents;
using AD.OpenXml.Html;
using JetBrains.Annotations;

namespace AD.OpenXml.Tests
{
    [UsedImplicitly]
    public static class Program
    {
        [UsedImplicitly]
        public static void Main()
        {
            // Declare working directory
            //const string path = @"c:\users\adren\desktop\508 work\";
            //const string path = @"g:\data\austin d\508 programming\508 work\";
            const string workingDirectory = "c:\\users\\adren\\desktop\\otap 2016\\ch2";
            //const string workingDirectory = "g:\\data\\austin d\\508 programming\\otap 2016\\ch2";

            // Create result file
            DocxFilePath result = DocxFilePath.Create($"{workingDirectory}\\output\\OTAP_2016_v1_0.docx", true);

            // Add footnotes file
            result.AddFootnotes();

            DocxFilePath[] files =
                Directory.GetFiles(workingDirectory, "*.docx", SearchOption.TopDirectoryOnly)
                         .OrderBy(
                             x =>
                             {
                                 string a = Path.GetFileName(x)?
                                                .TakeWhile(y => char.IsNumber(y) || char.IsPunctuation(y))
                                                .Where(char.IsNumber)
                                                .Aggregate(default(string), (current, next) => current + next)
                                                .Split('-')
                                                .FirstOrDefault();
                                 return double.Parse(a ?? "0");
                             })
                         .Select(x => (DocxFilePath) x)
                         .ToArray();

            OpenXmlContainer container = new OpenXmlContainer(result);

            foreach (DocxFilePath file in files)
            {
                container.MergeDocuments(file, result);
                container.Save(result);
            }

            // Create custom styles
            result.AddStyles();

            // Add headers
            result.AddHeaders("The Year in Trade 2016");

            // Add footers
            result.AddFooters();

            // Set all chart objects inline
            result.PositionChartsInline();

            // Set the inner positions of chart objects
            result.PositionChartsInner();

            // Set the outer positions of chart objects
            result.PositionChartsOuter();

            // Set the style of bar chart objects
            result.ModifyBarChartStyles();

            // Set the style of line chart objects
            result.ModifyLineChartStyles();

            // Remove duplicate section properties
            result.RemoveDuplicateSectionProperties();

            // Write document.xml to XML file
            XmlFilePath xml = XmlFilePath.Create($"{workingDirectory}\\TestWordDocument_out.xml");
            result.ReadAsXml().Elements().WriteXml(xml);

            // Write document.xml to HTML file
            HtmlFilePath html = HtmlFilePath.Create($"{workingDirectory}\\TestWordDocument_out.html");
            result.ReadAsXml().ProcessHtml().WriteHtml(html);
        }
    }
}