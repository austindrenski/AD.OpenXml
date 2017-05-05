using AD.OpenXml.Visitors;
using JetBrains.Annotations;

namespace AD.OpenXml.Visits
{
    /// <summary>
    /// 
    /// </summary>
    [PublicAPI]
    public interface IVisit
    {
        /// <summary>
        /// 
        /// </summary>
        OpenXmlVisitor Result { get; }
    }
}
