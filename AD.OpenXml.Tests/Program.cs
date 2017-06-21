using System;
using System.IO;
using System.Linq;
using AD.IO;
using AD.OpenXml.Documents;
using AD.OpenXml.Visitors;
using JetBrains.Annotations;

namespace AD.OpenXml.Tests
{
    [UsedImplicitly]
    public static class Program
    {
        [UsedImplicitly]
        public static void Main()
        {
            const string cberaTest = "g:\\data\\austin d\\cbera test";
            ProcessChapter("1_0", $"{cberaTest}\\_report");
            Console.ReadLine();

            // Declare working directory
            const string workingDirectory = "z:\\records\\operations\\economics\\sec 332\\active cases\\otap 2016\\draft report\\senior checkoff";

            // Declare version
            const string version = "4_0";

            //Process chapters
            ProcessChapter(version, $"{workingDirectory}\\ch0");
            ProcessChapter(version, $"{workingDirectory}\\ch1");
            ProcessChapter(version, $"{workingDirectory}\\ch2");
            ProcessChapter(version, $"{workingDirectory}\\ch3");
            ProcessChapter(version, $"{workingDirectory}\\ch4");
            ProcessChapter(version, $"{workingDirectory}\\ch5");
            ProcessChapter(version, $"{workingDirectory}\\ch6");
            ProcessChapter(version, $"{workingDirectory}\\ch7");

            // Copy new files into report folder
            foreach (string chapter in new string[] { "ch0", "ch1", "ch2", "ch3", "ch4", "ch5", "ch6", "ch7" })
            {
                Console.WriteLine(
                    Directory.GetFiles($"{workingDirectory}\\{chapter}\\_output", "*.docx", SearchOption.TopDirectoryOnly)
                             .Where(x => !x.Contains('~'))
                             .OrderByDescending(x => x.ParseLong())
                             .First());

                File.Copy(
                    Directory.GetFiles($"{workingDirectory}\\{chapter}\\_output", "*.docx", SearchOption.TopDirectoryOnly)
                             .Where(x => !x.Contains('~'))
                             .OrderByDescending(x => x.ParseLong())
                             .First(),
                    $"{workingDirectory}\\_report\\{chapter.ParseInt()} - {Path.GetFileName(chapter)}.docx",
                    true);
            }

            //Process report
            ProcessChapter(version, $"{workingDirectory}\\_report");

            ////Delete old files in report folder
            //foreach (string section in Directory.GetFiles($"{workingDirectory}\\_report", "*.docx", SearchOption.TopDirectoryOnly))
            //{
            //    File.Delete(section);
            //}
        }

        private static bool FilePredicate(string path)
        {
            if (path is null)
            {
                return false;
            }
            if (path.Contains('~'))
            {
                Console.WriteLine($"{DateTime.Now}: Skipping file '{path}'; unexpected character in path.");
                return false;
            }
            // ReSharper disable once InvertIf
            if (!char.IsNumber(Path.GetFileName(path).FirstOrDefault()))
            {
                Console.WriteLine($"{DateTime.Now}: Skipping file '{path}'; file name must start with number.");
                return false;
            }
            return true;
        }

        private static double? OrderPredicate(string path)
        {
            return
                Path.GetFileNameWithoutExtension(path ?? string.Empty)
                    .TakeWhile(y => char.IsNumber(y) || char.IsPunctuation(y))
                    .ParseDouble();
        }

        private static void ProcessChapter(string version, string workingDirectory)
        {
            // Create output directory
            Directory.CreateDirectory($"{workingDirectory}\\_output");

            // Locate the component files
            DocxFilePath[] files =
                Directory.GetFiles(workingDirectory, "*.docx", SearchOption.TopDirectoryOnly)
                         .Where(FilePredicate)
                         .OrderBy(OrderPredicate)
                         .Select(x => (DocxFilePath) x)
                         .ToArray();

            // Create output file
            DocxFilePath output = DocxFilePath.Create($"{workingDirectory}\\_output\\OTAP_2016_v_{version}.docx", true);
            
            // Create a ReportVisitor based on the result path and visit the component doucments.
            IOpenXmlVisitor visitor = new ReportVisitor(output).VisitAndFold(files);

            // Save the visitor results to result path.
            visitor.Save(output);

            // Add headers
            output.AddHeaders("The Year in Trade 2016");

            // Add footers
            output.AddFooters();

            // Set all chart objects inline
            output.PositionChartsInline();

            // Set the inner positions of chart objects
            output.PositionChartsInner();

            // Set the outer positions of chart objects
            output.PositionChartsOuter();

            // Set the style of bar chart objects
            output.ModifyBarChartStyles();

            // Set the style of pie chart objects
            output.ModifyPieChartStyles();

            // Set the style of line chart objects
            output.ModifyLineChartStyles();
            
            // Set the style of area chart objects
            output.ModifyAreaChartStyles();

            // Remove duplicate section properties
            //output.RemoveDuplicateSectionProperties();

            // Write document.xml to XML file
            //XmlFilePath xml = XmlFilePath.Create($"{workingDirectory}\\_output\\OTAP_2016_v_{version}.xml");
            //result.ReadAsXml().Elements().WriteXml(xml);

            // Write document.xml to HTML file
            //HtmlFilePath html = HtmlFilePath.Create($"{workingDirectory}\\_output\\OTAP_2016_v_{version}.html");
            //result.ReadAsXml().ProcessHtml().WriteHtml(html);
        }
    }
}