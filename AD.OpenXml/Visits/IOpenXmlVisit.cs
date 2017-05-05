using AD.OpenXml.Visitors;
using JetBrains.Annotations;

namespace AD.OpenXml.Visits
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
