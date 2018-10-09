using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using AD.IO.Paths;
using AD.OpenXml;
using AD.OpenXml.Documents;
using AD.OpenXml.Markdown;
using AD.OpenXml.Structures;
using JetBrains.Annotations;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

namespace CompilerAPI.Controllers
{
    /// <inheritdoc />
    /// <summary>
    /// Provides endpoints to format and normalize Word documents.
    /// </summary>
    [PublicAPI]
    [FormatFilter]
    [Route("[controller]")]
    [ApiVersion("2.0")]
    [ApiVersion("1.0", Deprecated = true)]
    public class UploadController : Controller
    {
        [NotNull] const string MicrosoftWordDocument =
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

        /// <summary>
        /// Returns the webpage with an upload form for documents.
        /// </summary>
        /// <returns>
        /// The index razor view.
        /// </returns>
        [HttpGet("")]
        [ProducesResponseType(StatusCodes.Status200OK)]
        public ViewResult Index() => View();

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
        /// <exception cref="ArgumentNullException"><paramref name="files"/></exception>
        [HttpPost("")]
        [ProducesResponseType(StatusCodes.Status200OK)]
        [ProducesResponseType(StatusCodes.Status400BadRequest)]
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

            if (uploadedFiles.Any(x => x.ContentType != MicrosoftWordDocument.ToString() &&
                                       x.FileName.EndsWith(".docx", StringComparison.OrdinalIgnoreCase) &&
                                       x.FileName.EndsWith(".md", StringComparison.OrdinalIgnoreCase)))
                return BadRequest("Invalid file format.");

            if (stylesheet is IFormFile s && !s.FileName.EndsWith(".css"))
                return BadRequest($"Invalid stylesheet:{stylesheet.FileName}.");

            Queue<Package> packagesQueue = new Queue<Package>(uploadedFiles.Length);

            foreach (IFormFile file in uploadedFiles)
            {
                if (file.FileName.EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
                {
                    packagesQueue.Enqueue(Package.Open(file.OpenReadStream()));
                }
                else if (file.FileName.EndsWith(".md", StringComparison.OrdinalIgnoreCase))
                {
                    MDocument mDocument = new MDocument();
                    using (StreamReader reader = new StreamReader(file.OpenReadStream()))
                    {
                        MarkdownVisitor visitor = new MarkdownVisitor();
                        ReadOnlySpan<char> span;
                        while ((span = reader.ReadLine()) != null)
                        {
                            mDocument.Append(visitor.Visit(in span));
                        }
                    }

                    Package package = DocxFilePath.Create().ToPackage(FileAccess.ReadWrite);

                    mDocument.ToOpenXml().WriteTo(package.GetPart(Document.PartUri));

                    packagesQueue.Enqueue(package);
                }
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
                {
                    return File(output.ToStream(), MicrosoftWordDocument, $"{title ?? "result"}.docx");
                }
                case "html":
                {
                    string styles = stylesheet is null ? null : new StreamReader(stylesheet.OpenReadStream()).ReadToEnd();
                    OpenXmlPackageVisitor ooxml = new OpenXmlPackageVisitor(output);
                    HtmlVisitor html = new HtmlVisitor(ooxml.Document.ChartReferences, ooxml.Document.ImageReferences);
                    XObject htmlResult = html.Visit(ooxml.Document.Content, ooxml.Footnotes.Content, title, stylesheetUrl, styles);
                    return File(Encoding.UTF8.GetBytes(htmlResult.ToString()), "text/html", $"{title ?? "result"}.html");
                }
                case "xml":
                {
                    OpenXmlPackageVisitor xml = new OpenXmlPackageVisitor(output);
                    XElement xmlResult = xml.Document.Content;
                    return File(Encoding.UTF8.GetBytes(xmlResult.ToString()), "application/xml", $"{title ?? "result"}.xml");
                }
                default:
                {
                    return BadRequest(ModelState);
                }
            }
        }

        [Pure]
        [NotNull]
        static Package Process(
            [NotNull] [ItemNotNull] IEnumerable<Package> packages,
            [NotNull] string title,
            [NotNull] string publisher,
            [NotNull] string website)
            => OpenXmlPackageVisitor.Visit(packages)
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