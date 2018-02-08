using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
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
        /// <returns>
        ///
        /// </returns>
        [NotNull]
        [HttpGet]
        public IActionResult Index()
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
        [ItemNotNull]
        public async Task<IActionResult> Index([NotNull] [ItemNotNull] IEnumerable<IFormFile> files)
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

            if (uploadedFiles.Any(x => x.ContentType != _microsoftWordDocument.ToString()))
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

            output.Seek(0, SeekOrigin.Begin);

            return new FileStreamResult(output, _microsoftWordDocument);
        }

        [Pure]
        [NotNull]
        [ItemNotNull]
        private static async Task<MemoryStream> Process([NotNull] [ItemNotNull] IEnumerable<MemoryStream> files, [NotNull] string reportTitle)
        {
            if (files is null)
            {
                throw new ArgumentNullException(nameof(files));
            }
            if (reportTitle is null)
            {
                throw new ArgumentNullException(nameof(reportTitle));
            }

            return
                await new ReportVisitor()
                    .VisitAndFold(files)
                    .Save()
                    .AddHeaders(reportTitle)
                    .AddFooters()
                    .PositionChartsInline()
                    .PositionChartsInner()
                    .PositionChartsOuter()
                    .ModifyBarChartStyles()
                    .ModifyPieChartStyles()
                    .ModifyLineChartStyles()
                    .ModifyAreaChartStyles();
        }
    }
}