using JetBrains.Annotations;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

namespace CompilerAPI.Controllers
{
    /// <inheritdoc />
    /// <summary>
    /// Provides endpoints for the web frontend of the Reports API.
    /// </summary>
    [PublicAPI]
    [FormatFilter]
    [Route("")]
    [ApiController]
    [ApiVersion("2.0")]
    [ApiVersion("1.0", Deprecated = true)]
    public class HomeController : Controller
    {
        /// <summary>
        /// The endpoint for the web frontend of the Reports API.
        /// </summary>
        /// <returns>
        /// An HTML page describing Reports API.
        /// </returns>
        [HttpGet("")]
        [ProducesResponseType(StatusCodes.Status307TemporaryRedirect)]
        public IActionResult Index() => RedirectToAction("Index", "Upload");

        /// <summary>
        /// The endpoint for guidance on how to format documents for the Reports API.
        /// </summary>
        /// <returns>
        /// An HTML page describing how to format documents for the Reports API.
        /// </returns>
        [HttpGet("[action]")]
        [ProducesResponseType(StatusCodes.Status200OK)]
        public ViewResult Guidance() => View();

        /// <summary>
        /// The endpoint for an overview of the Reports API process.
        /// </summary>
        /// <returns>
        /// An HTML page describing the Reports API process.
        /// </returns>
        [HttpGet("[action]")]
        [ProducesResponseType(StatusCodes.Status200OK)]
        public ViewResult Process() => View();

        /// <summary>
        /// The endpoint for the guidance on using the Reports API for report production.
        /// </summary>
        /// <returns>
        /// An HTML page describing the Reports API production processes.
        /// </returns>
        [HttpGet("[action]")]
        [ProducesResponseType(StatusCodes.Status200OK)]
        public ViewResult Production() => View();
    }
}