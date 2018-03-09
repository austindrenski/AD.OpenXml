using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using AD.OpenXml.Documents;
using AD.OpenXml.Html;
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
    /// Provides HTTP endpoints to submit and format Word documents.
    /// </summary>
    [PublicAPI]
    [ApiVersion("1.0")]
    [Route("[controller]")]
    public class UploadController : Controller
    {
        private static MediaTypeHeaderValue _microsoftWordDocument = new MediaTypeHeaderValue("application/vnd.openxmlformats-officedocument.wordprocessingml.document");

        /// <summary>
        /// Returns the webpage with an upload form for documents.
        /// </summary>
        /// <returns>
        /// The index razor view.
        /// </returns>
        [NotNull]
        [HttpGet("")]
        public IActionResult Index()
        {
            return View();
        }

        /// <summary>
        /// Receives file uploads from the user.
        /// </summary>
        /// <param name="files">
        /// The collection of files submitted by POST request.
        /// </param>
        /// <param name="format">
        /// The format to produce
        /// </param>
        /// <param name="title">
        /// The title of the document to be returned.
        /// </param>
        /// <param name="publisher">
        /// The name of the publisher for the document to be returned.
        /// </param>
        /// <param name="website">
        /// The website of the publisher.
        /// </param>
        /// <returns>
        /// The combined and formatted document.
        /// </returns>
        /// <exception cref="ArgumentNullException"/>
        [NotNull]
        [HttpPost("")]
        [ItemNotNull]
        public async Task<IActionResult> Index([NotNull] [ItemNotNull] IEnumerable<IFormFile> files, [CanBeNull] string format, [CanBeNull] string title, [CanBeNull] string publisher, [CanBeNull] string website)
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

            MemoryStream output =
                await Process(
                    documentQueue,
                    title ?? "[REPORT TITLE]",
                    publisher ?? "[PUBLISHER]",
                    website ?? "[PUBLISHER WEBSITE]");

            output.Seek(0, SeekOrigin.Begin);

            if (format != "html")
            {
                return new FileStreamResult(output, _microsoftWordDocument);
            }

            ReportVisitor visitor = new ReportVisitor(output);

            return
                new ContentResult
                {
                    Content = visitor.Document.Elements().Single().BodyToHtml().ToString(),
                    ContentType = "text/html",
                    StatusCode = 200
                };
        }

        [Pure]
        [NotNull]
        [ItemNotNull]
        private static async Task<MemoryStream> Process([NotNull] [ItemNotNull] IEnumerable<MemoryStream> files, [NotNull] string title, [NotNull] string publisher, [NotNull] string website)
        {
            if (files is null)
            {
                throw new ArgumentNullException(nameof(files));
            }
            
            if (title is null)
            {
                throw new ArgumentNullException(nameof(title));
            }
            
            if (publisher is null)
            {
                throw new ArgumentNullException(nameof(publisher));
            }
            
            if (website is null)
            {
                throw new ArgumentNullException(nameof(website));
            }

            if (publisher is null)
            {
                throw new ArgumentNullException(nameof(publisher));
            }

            if (website is null)
            {
                throw new ArgumentNullException(nameof(website));
            }

            return
                await new ReportVisitor()
                      .VisitAndFold(files)
                      .Save()
                      .AddHeaders(title)
                      .AddFooters(publisher, website)
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