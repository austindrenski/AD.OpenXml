using JetBrains.Annotations;
using Microsoft.AspNetCore.Mvc;

namespace EditorAPI.Controllers
{
    /// <summary>
    ///
    /// </summary>
    [Route("")]
    public class EditorController : Controller
    {
        /// <summary>
        ///
        /// </summary>
        /// <returns>
        ///
        /// </returns>
        [Pure]
        [NotNull]
        [HttpGet("")]
        public IActionResult Index() => View();
    }
}