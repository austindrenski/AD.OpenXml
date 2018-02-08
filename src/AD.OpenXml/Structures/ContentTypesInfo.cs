using System.Xml.Linq;
using JetBrains.Annotations;

namespace AD.OpenXml.Structures
{
    // TODO: document ContentTypesInfo file.
    /// <summary>
    ///
    /// </summary>
    [PublicAPI]
    public static class ContentTypesInfo
    {
        /// <summary>
        ///
        /// </summary>
        [NotNull] public const string Path = "[Content_Types].xml";

        /// <summary>
        ///
        /// </summary>
        [NotNull] public static readonly XNamespace Namespace = "http://schemas.openxmlformats.org/package/2006/content-types";

        /// <summary>
        ///
        /// </summary>
        [NotNull] public static readonly XName Root = "Types";

        /// <summary>
        ///
        /// </summary>
        [PublicAPI]
        public static class Attributes
        {
            /// <summary>
            ///
            /// </summary>
            [NotNull] public static readonly XName ContentType = "ContentType";

            /// <summary>
            ///
            /// </summary>
            [NotNull] public static readonly XName Extension = "Extension";

            /// <summary>
            ///
            /// </summary>
            [NotNull] public static readonly XName PartName = "PartName";
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
            [NotNull] public static readonly XName Default = Namespace + "Default";

            /// <summary>
            ///
            /// </summary>
            [NotNull] public static readonly XName Override = Namespace + "Override";
        }
    }
}