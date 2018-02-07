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
    [Route("[controller]/[action]")]
    public class Home : Controller
    {
        /// <summary>
        ///
        /// </summary>
        /// <returns></returns>
        [NotNull]
        [HttpGet]
        public IActionResult Index()
        {
            return View();
        }
    }
}