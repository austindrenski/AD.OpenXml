using System;
using System.Xml.Linq;
using JetBrains.Annotations;

namespace AD.OpenXml.Structures
{
    // TODO: Document FootnotesRelsInfo file.
    /// <summary>
    ///
    /// </summary>
    [PublicAPI]
    public static class FootnotesRelsInfo
    {
        /// <summary>
        ///
        /// </summary>
        [NotNull]
        public static Uri PartName => new Uri("/word/_rels/footnotes.xml.rels", UriKind.Relative);

        /// <summary>
        ///
        /// </summary>
        [NotNull] public static readonly XNamespace Namespace = "http://schemas.openxmlformats.org/package/2006/relationships";

        /// <summary>
        ///
        /// </summary>
        [NotNull] public static readonly XName Root = "Relationships";

        /// <summary>
        ///
        /// </summary>
        [PublicAPI]
        public static class Attributes
        {
            /// <summary>
            ///
            /// </summary>
            [NotNull] public static readonly XName Id = "Id";

            /// <summary>
            ///
            /// </summary>
            [NotNull] public static readonly XName Type = "Type";

            /// <summary>
            ///
            /// </summary>
            [NotNull] public static readonly XName Target = "Target";

            /// <summary>
            ///
            /// </summary>
            [NotNull] public static readonly XName TargetMode = "TargetMode";
        }

        /// <summary>
        ///
        /// </summary>
        [PublicAPI]
        public static class Elements
        {
            /// <summary>
            ///
            /// </summary>
            [NotNull] public static readonly XName Relationship = Namespace + "Relationship";
        }
    }
}