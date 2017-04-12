using AD.IO;
using AD.OpenXml.Documents;
using JetBrains.Annotations;

namespace AD.OpenXml.Tests
{
    [UsedImplicitly]
    public static class DemoProgram
    {
        [UsedImplicitly]
        public static void Demo()
        {
            // Declare working directory
            const string path = @"G:\data\austin d\508 programming\508 presentation to oars-ocio (1-25-17)\";

            // Declare template and source file
            DocxFilePath source = DocxFilePath.Create(path + "Introduction example.docx");

            // Create result file from template
            DocxFilePath result = DocxFilePath.Create(path + "Result - Introduction example.docx", true);

            //// Declare template and source file
            //DocxFilePath source = DocxFilePath.Create(path + "Testing examples.docx");

            //// Create result file from template
            //DocxFilePath result = DocxFilePath.Create(path + "Result - Testing example.docx", true);

            // Clean document.xml
            source.WriteInto(result, "word/document.xml");

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
        }

        //private static void Demo2()
        //{
        //    // Declare working directory
        //    //const string path = @"c:\users\adren\desktop\508 work\";
        //    //const string path = @"g:\data\austin d\508 programming\508 work\";
        //    const string path = @"g:\data\austin d\508 programming\OTAP 2016\";

        //    // Declare source file
        //    DocxFilePath source = DocxFilePath.Create(path + "TestWordDocument.docx");

        //    // Create result file
        //    DocxFilePath result = DocxFilePath.Create(path + "TestWordDocument_out.docx", true);

        //    // Clean document.xml
        //    result.Process508From(source);

        //    // Test AltChunk implementation
        //    result.ReadAsXml().Elements().First().AddAltChunk(result, source).WriteInto(result, "word/document.xml");

        //    // Create custom styles
        //    result.AddStyles();

        //    // Add headers
        //    result.AddHeaders("Year in Trade");

        //    // Add footers
        //    result.AddFooters();

        //    // Set all chart objects inline
        //    result.PositionChartsInline();

        //    // Set the inner positions of chart objects
        //    result.PositionChartsInner();

        //    // Set the outer positions of chart objects
        //    result.PositionChartsOuter();

        //    // Set the style of bar chart objects
        //    result.ModifyBarChartStyles();

        //    // Set the style of line chart objects
        //    result.ModifyLineChartStyles();

        //    // Write document.xml to XML file
        //    XmlFilePath xml = XmlFilePath.Create(path + "TestWordDocument_out.xml");
        //    result.ReadAsXml().Elements().WriteXml(xml);

        //    // Write document.xml to HTML file
        //    HtmlFilePath html = HtmlFilePath.Create(path + "TestWordDocument_out.html");
        //    result.ReadAsXml().ProcessHtml().WriteHtml(html);
        //}
    }
}