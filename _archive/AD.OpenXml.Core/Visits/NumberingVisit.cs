using System.Xml.Linq;
using AD.OpenXml.Core.Visitors;

namespace AD.OpenXml.Core.Visits
{
    /// <summary>
    /// 
    /// </summary>
    [PublicAPI]
    public sealed class NumberingVisit : IOpenXmlVisit
    {
        [NotNull]
        private static readonly XNamespace T = XNamespaces.OpenXmlPackageContentTypes;

        [NotNull]
        private static readonly XNamespace P = XNamespaces.OpenXmlPackageRelationships;

        [NotNull]
        private static readonly XNamespace W = XNamespaces.OpenXmlWordprocessingmlMain;

        /// <summary>
        /// 
        /// </summary>
        public IOpenXmlVisitor Result { get; }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="subject"></param>
        public NumberingVisit(IOpenXmlVisitor subject)
        {
            XElement numbering = Execute();

            Result =
                new OpenXmlVisitor(
                    subject.ContentTypes,
                    subject.Document,
                    subject.DocumentRelations,
                    subject.Footnotes,
                    subject.FootnoteRelations,
                    subject.Styles,
                    numbering,
                    subject.Charts);
        }

        [Pure]
        private static XElement Execute()
        {
            XElement numbering =
                XElement.Parse(Resources.Numbering);         

            return numbering;
        }
    }
}
