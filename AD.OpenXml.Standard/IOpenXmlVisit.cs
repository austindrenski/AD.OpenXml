using JetBrains.Annotations;

namespace AD.OpenXml
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
