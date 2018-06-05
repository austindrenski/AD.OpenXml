using System;
using System.IO;
using System.Reflection;
using System.Text;
using System.Xml.Linq;
using AD.Xml;
using JetBrains.Annotations;

namespace AD.OpenXml.Visits
{
    /// <summary>
    ///
    /// </summary>
    [PublicAPI]
    public static class NumberingVisit
    {
        [NotNull] private static readonly XNamespace T = XNamespaces.OpenXmlPackageContentTypes;
        [NotNull] private static readonly XNamespace P = XNamespaces.OpenXmlPackageRelationships;
        [NotNull] private static readonly XNamespace W = XNamespaces.OpenXmlWordprocessingmlMain;
        [NotNull] private static readonly XElement Numbering;

        /// <summary>
        ///
        /// </summary>
        /// <exception cref="FileNotFoundException" />
        static NumberingVisit()
        {
            Assembly assembly = typeof(NumberingVisit).GetTypeInfo().Assembly;

            using (StreamReader reader =
                new StreamReader(
                    assembly.GetManifestResourceStream("AD.OpenXml.Styles.Numbering.xml")
                    ?? throw new FileNotFoundException("AD.OpenXml.Styles.Numbering.xml"),
                    Encoding.UTF8))
            {
                Numbering = XElement.Parse(reader.ReadToEnd());
            }
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="subject"></param>
        /// <exception cref="ArgumentNullException" />
        [Pure]
        [NotNull]
        public static OpenXmlPackageVisitor VisitNumbering([NotNull] this OpenXmlPackageVisitor subject)
        {
            if (subject is null)
                throw new ArgumentNullException(nameof(subject));

            XElement numbering = Numbering.Clone();

            return subject.With(numbering: numbering);
        }
    }
}