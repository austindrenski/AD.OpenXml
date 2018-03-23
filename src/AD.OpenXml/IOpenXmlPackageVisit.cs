using JetBrains.Annotations;

namespace AD.OpenXml
{
    /// <summary>
    /// Represents a visit to an <see cref="OpenXmlPackageVisitor"/>.
    /// </summary>
    [PublicAPI]
    public interface IOpenXmlPackageVisit
    {
        /// <summary>
        /// The result of the visit.
        /// </summary>
        OpenXmlPackageVisitor Result { get; }
    }
}