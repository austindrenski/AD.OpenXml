using JetBrains.Annotations;
using Microsoft.AspNetCore.Mvc;

namespace CompilerAPI.Controllers
{
    /// <inheritdoc />
    /// <summary>
    ///
    /// </summary>
    [PublicAPI]
    [ApiVersion("1.0")]
    public class HomeController : Controller
    {
        /// <summary>
        ///
        /// </summary>
        /// <returns>
        ///
        /// </returns>
        [Pure]
        [NotNull]
        [HttpGet]
        public IActionResult Index() => RedirectToAction("Index", "Upload");
    }
}