using System;
using System.IO;
using System.IO.Packaging;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using JetBrains.Annotations;

namespace AD.OpenXml
{
    /// <summary>
    /// Provides extension methods to write an <see cref="XElement"/> to a <see cref="PackagePart"/>.
    /// </summary>
    [PublicAPI]
    public static class WriteToPart
    {
        [NotNull] private static readonly XmlWriterSettings XmlWriterSettings =
            new XmlWriterSettings
            {
                Async = false,
                DoNotEscapeUriAttributes = false,
                CheckCharacters = true,
                CloseOutput = true,
                ConformanceLevel = ConformanceLevel.Document,
                Encoding = Encoding.UTF8,
                Indent = false,
                IndentChars = "  ",
                NamespaceHandling = NamespaceHandling.OmitDuplicates,
                NewLineChars = Environment.NewLine,
                NewLineHandling = NewLineHandling.None,
                NewLineOnAttributes = false,
                OmitXmlDeclaration = false,
                WriteEndDocumentOnClose = true
            };

        /// <summary>
        /// Writes the <paramref name="node"/> to the <paramref name="part" />.
        /// </summary>
        /// <param name="node">The node to write.</param>
        /// <param name="part">The part to which the element is written.</param>
        /// <exception cref="ArgumentNullException" />
        public static void WriteTo([NotNull] this XNode node, [NotNull] PackagePart part)
        {
            if (node is null)
                throw new ArgumentNullException(nameof(node));
            if (part is null)
                throw new ArgumentNullException(nameof(part));

            using (XmlWriter xml = XmlWriter.Create(part.GetStream(FileMode.Create), XmlWriterSettings))
            {
                node.WriteTo(xml);
            }
        }

        /// <summary>
        /// Reads the <see cref="XElement"/> from the <paramref name="part" />.
        /// </summary>
        /// <param name="part">The part from which the element is read.</param>
        /// <returns>
        /// The <see cref="XElement"/> of the specified part and relationship.
        /// </returns>
        /// <exception cref="ArgumentNullException" />
        [Pure]
        public static XElement ReadXml([NotNull] this PackagePart part)
        {
            if (part is null)
                throw new ArgumentNullException(nameof(part));

            using (Stream stream = part.GetStream())
            {
                return XElement.Load(stream);
            }
        }
    }
}