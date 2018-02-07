using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using AD.IO.Paths;
using AD.OpenXml;
using AD.OpenXml.Documents;
using AD.OpenXml.Visitors;
using JetBrains.Annotations;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Net.Http.Headers;

namespace CompilerAPI.Controllers
{
    // TODO: document UploadController
    /// <inheritdoc />
    /// <summary>
    ///
    /// </summary>
    [PublicAPI]
    [ApiVersion("1.0")]
    [Route("[controller]/[action]")]
    public class UploadController : Controller
    {
        private static MediaTypeHeaderValue _microsoftWordDocument = new MediaTypeHeaderValue("application/vnd.openxmlformats-officedocument.wordprocessingml.document");

        /// <summary>
        ///
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        public IActionResult Post()
        {
            return View();
        }

        /// <summary>
        ///
        /// </summary>
        /// <returns>
        ///
        /// </returns>
        [NotNull]
        [HttpGet]
        public IActionResult StreamTest()
        {
            return View();
        }

        /// <summary>
        ///
        /// </summary>
        /// <returns>
        ///
        /// </returns>
        [NotNull]
        [HttpPost]
        public async Task<IActionResult> StreamTest([NotNull] [ItemNotNull] IEnumerable<IFormFile> files)
        {
            if (files is null)
            {
                throw new ArgumentNullException(nameof(files));
            }

            IFormFile[] uploadedFiles = files.ToArray();

            if (uploadedFiles.Length == 0)
            {
                return BadRequest("No files uploaded.");
            }

            if (uploadedFiles.Any(x => x.Length <= 0))
            {
                return BadRequest("Invalid file length.");
            }

            if (uploadedFiles.Any(x => !Path.GetExtension(x.FileName).Equals(".docx", StringComparison.OrdinalIgnoreCase)))
            {
                return BadRequest("Invalid file format.");
            }

            Queue<MemoryStream> documentQueue = new Queue<MemoryStream>(uploadedFiles.Length);

            foreach (IFormFile file in uploadedFiles)
            {
                MemoryStream memoryStream = new MemoryStream();
                await file.CopyToAsync(memoryStream);
                documentQueue.Enqueue(memoryStream);
            }

            MemoryStream output = await Process(documentQueue, "[REPORT TITLE HERE]");

            return new FileStreamResult(output, _microsoftWordDocument);
        }

        private static async Task<MemoryStream> Process(IEnumerable<MemoryStream> files, string reportTitle)
        {
            // Create a ReportVisitor based on the result path and visit the component doucments.
            IOpenXmlVisitor visitor = new ReportVisitor(DocxFilePath.Create($"{Path.GetTempFileName()}.docx")).VisitAndFold(files);

            // Save the visitor results to result path.
            // Add headers
            MemoryStream result =
                await visitor.Save()
                             .AddHeaders(reportTitle);

//            // Add footers
//            result.AddFooters();
//
//            // Set all chart objects inline
//            result.PositionChartsInline();
//
//            // Set the inner positions of chart objects
//            result.PositionChartsInner();
//
//            // Set the outer positions of chart objects
//            result.PositionChartsOuter();
//
//            // Set the style of bar chart objects
//            result.ModifyBarChartStyles();
//
//            // Set the style of pie chart objects
//            result.ModifyPieChartStyles();
//
//            // Set the style of line chart objects
//            result.ModifyLineChartStyles();
//
//            // Set the style of area chart objects
//            result.ModifyAreaChartStyles();

            return result;
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="files"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentNullException"/>
        [HttpPost]
        public async Task<IActionResult> Post([NotNull] IEnumerable<IFormFile> files)
        {
            if (files is null)
            {
                throw new ArgumentNullException(nameof(files));
            }

            IFormFile[] uploadedFiles = files.ToArray();

            if (uploadedFiles.Length == 0)
            {
                return BadRequest("No files uploaded.");
            }
            if (uploadedFiles.Any(x => x.Length <= 0))
            {
                return BadRequest("Invalid file length.");
            }
            if (uploadedFiles.Any(x => !Path.GetExtension(x.FileName).Equals(".docx", StringComparison.OrdinalIgnoreCase)))
            {
                return BadRequest("Invalid file format.");
            }

            Queue<DocxFilePath> inputQueue = new Queue<DocxFilePath>(uploadedFiles.Length);

            foreach (IFormFile file in uploadedFiles)
            {
                DocxFilePath input = DocxFilePath.Create(Path.ChangeExtension(Path.GetTempFileName(), "docx"), true);

                using (FileStream fileStream = new FileStream(input, FileMode.Open))
                {
                    await file.CopyToAsync(fileStream);
                }

                inputQueue.Enqueue(input);
            }

            DocxFilePath output = DocxFilePath.Create(Path.ChangeExtension(Path.GetTempFileName(), "docx"), true);

            Process(inputQueue, output, "[REPORT TITLE HERE]");

            return new FileStreamResult(new FileStream(output, FileMode.Open), _microsoftWordDocument);
        }

        private static void Process(IEnumerable<DocxFilePath> files, DocxFilePath output, string reportTitle)
        {
            // Create a ReportVisitor based on the result path and visit the component doucments.
            IOpenXmlVisitor visitor = new ReportVisitor(output).VisitAndFold(files);

            // Save the visitor results to result path.
            visitor.Save(output);

            // Add headers
            output.AddHeaders(reportTitle);

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
        }
    }
}