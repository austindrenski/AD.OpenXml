using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
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
            const string workingDirectory = "g:\\data\\austin d\\508 programming\\otap 2016\\ch5";

            // Create result file
            DocxFilePath result = DocxFilePath.Create($"{workingDirectory}\\output\\OTAP_2016_v1_0.docx", true);

            foreach (string file in Directory.GetFiles(workingDirectory).Where(x => x.EndsWith(".docx")))
            {
                Combine(file, result);
            }

            // Create custom styles
            result.AddStyles();

            // Add headers
            result.AddHeaders("Year in Trade");

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

            // Write document.xml to XML file
            XmlFilePath xml = XmlFilePath.Create($"{workingDirectory}\\TestWordDocument_out.xml");
            result.ReadAsXml().Elements().WriteXml(xml);

            // Write document.xml to HTML file
            HtmlFilePath html = HtmlFilePath.Create($"{workingDirectory}\\TestWordDocument_out.html");
            result.ReadAsXml().ProcessHtml().WriteHtml(html);
        }

        private static void Combine(DocxFilePath source , DocxFilePath result)
        {
            DocxFilePath tempSource = DocxFilePath.Create($"{source}_temp.docx", true);

            tempSource.Process508From(source);

            XElement sourceDocument = tempSource.ReadAsXml();
            XElement resultDocument = result.ReadAsXml();

            XElement sourceBody =
                sourceDocument.Elements()
                              .Single(x => x.Name.LocalName.Equals("body"));

            IEnumerable<XElement> sourceContent =
                sourceBody.Elements();

            resultDocument.Elements().First().Add(sourceContent);

            resultDocument.WriteInto(result, "word/document.xml");

            File.Delete(tempSource);
        }
    }
}