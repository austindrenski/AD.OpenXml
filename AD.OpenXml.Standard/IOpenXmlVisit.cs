using System.Linq;
using JetBrains.Annotations;

namespace AD.OpenXml.Standard
{
    /// <summary>
    /// 
    /// </summary>
    [PublicAPI]
    public interface IOpenXmlVisit
    {
        /// <summary>
        /// 
        /// </summary>
        IOpenXmlVisitor Result { get; }
    }
}
