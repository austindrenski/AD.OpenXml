using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;
using AD.OpenXml;
using AD.OpenXml.Documents;
using JetBrains.Annotations;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Net.Http.Headers;

namespace CompilerAPI.Controllers
{
    /// <inheritdoc />
    /// <summary>
    /// Provides HTTP endpoints to submit and format Word documents.
    /// </summary>
    [PublicAPI]
    [FormatFilter]
    [ApiVersion("1.0")]
    public class UploadController : Controller
    {
        private static MediaTypeHeaderValue _microsoftWordDocument =
            new MediaTypeHeaderValue("application/vnd.openxmlformats-officedocument.wordprocessingml.document");

        /// <summary>
        /// Returns the webpage with an upload form for documents.
        /// </summary>
        /// <returns>
        /// The index razor view.
        /// </returns>
        [NotNull]
        [HttpGet]
        public IActionResult Index() => View();

        /// <summary>
        /// Receives file uploads from the user.
        /// </summary>
        /// <param name="files">The collection of files submitted by POST request.</param>
        /// <param name="format">The format to produce.</param>
        /// <param name="title">The title of the document to be returned.</param>
        /// <param name="publisher">The name of the publisher for the document to be returned.</param>
        /// <param name="website">The website of the publisher.</param>
        /// <param name="stylesheetUrl">A style sheet reference link to be included on the HTML document.</param>
        /// <param name="stylesheet">A style sheet to embed in the HTML document.</param>
        /// <returns>
        /// The combined and formatted document.
        /// </returns>
        /// <exception cref="ArgumentNullException"/>
        [Pure]
        [NotNull]
        [HttpPost]
        public IActionResult Index(
            [NotNull] [ItemNotNull] IEnumerable<IFormFile> files,
            [CanBeNull] string format,
            [CanBeNull] string title,
            [CanBeNull] string publisher,
            [CanBeNull] string website,
            [CanBeNull] string stylesheetUrl,
            [CanBeNull] IFormFile stylesheet)
        {
            if (files is null)
                throw new ArgumentNullException(nameof(files));

            IFormFile[] uploadedFiles = files.ToArray();

            if (uploadedFiles.Length == 0)
                return BadRequest("No files uploaded.");

            if (uploadedFiles.Any(x => x.Length <= 0))
                return BadRequest("Invalid file length.");

            if (uploadedFiles.Any(x => x.ContentType != _microsoftWordDocument.ToString() &&
                                       x.FileName.EndsWith(".docx", StringComparison.OrdinalIgnoreCase) &&
                                       x.FileName.EndsWith(".md", StringComparison.OrdinalIgnoreCase)))
                return BadRequest("Invalid file format.");

            if (stylesheet is IFormFile s && !s.FileName.EndsWith(".css"))
                return BadRequest($"Invalid stylesheet:{stylesheet.FileName}.");

            Queue<Package> packagesQueue = new Queue<Package>(uploadedFiles.Length);

            foreach (IFormFile file in uploadedFiles)
            {
                if (file.FileName.EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
                    packagesQueue.Enqueue(Package.Open(file.OpenReadStream()));

//                else
//                {
//                    StringSegment markdown = new StreamReader(file.OpenReadStream()).ReadToEnd();
//                    MNode result = new MarkdownVisitor().Visit(in markdown);
//
//                    ZipArchive archive = new ZipArchive(DocxFilePath.Create(), ZipArchiveMode.Update);
//
//                    using (Stream stream = archive.GetEntry("word/document.xml")?.Open())
//                    {
//                        (result.ToOpenXml() as XElement)?.Save(stream);
//                    }
//
//                    documentQueue.Enqueue(archive);
//                }
            }

            Package output =
                Process(
                    packagesQueue,
                    title ?? "[REPORT TITLE]",
                    publisher ?? "[PUBLISHER]",
                    website ?? "[PUBLISHER WEBSITE]");

            foreach (Package package in packagesQueue)
            {
                package.Close();
            }

            switch (format)
            {
                case "docx":
                    return new FileStreamResult(output.ToStream(), _microsoftWordDocument);

                case "html":
                {
                    string styles = stylesheet is null ? null : new StreamReader(stylesheet.OpenReadStream()).ReadToEnd();
                    OpenXmlPackageVisitor ooxml = new OpenXmlPackageVisitor(output);
                    HtmlVisitor html = new HtmlVisitor(ooxml.Document.ChartReferences, ooxml.Document.ImageReferences);
                    XObject result = html.Visit(ooxml.Document.Content, ooxml.Footnotes.Content, title, stylesheetUrl, styles);
                    return Content(result.ToString(), "text/html");
                }
                case "xml":
                    return Content(new OpenXmlPackageVisitor(output).Document.Content.ToString(), "application/xml");

                default:
                    return BadRequest(ModelState);
            }
        }

        [Pure]
        [NotNull]
        private static Package Process(
            [NotNull] [ItemNotNull] IEnumerable<Package> packages,
            [NotNull] string title,
            [NotNull] string publisher,
            [NotNull] string website)
            => OpenXmlPackageVisitor
              .Visit(packages)
              .Package
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