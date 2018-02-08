using JetBrains.Annotations;

namespace AD.OpenXml.Structure
{
    // TODO: document DocumentRelsInfo file.
    /// <summary>
    ///
    /// </summary>
    [PublicAPI]
    public static class DocumentRelsInfo
    {
        /// <summary>
        ///
        /// </summary>
        [NotNull] public const string Path = "word/_rels/document.xml.rels";

        /// <summary>
        ///
        /// </summary>
        [NotNull] public const string Namespace = "http://schemas.openxmlformats.org/package/2006/relationships";

        /// <summary>
        ///
        /// </summary>
        [NotNull] public const string Root = "Relationships";

        /// <summary>
        ///
        /// </summary>
        public const bool IsNamespacePrefixed = false;

        /// <summary>
        ///
        /// </summary>
        [PublicAPI]
        public static class Attributes
        {
            /// <summary>
            ///
            /// </summary>
            [NotNull] public const string Id = "Id";

            /// <summary>
            ///
            /// </summary>
            [NotNull] public const string Type = "Type";

            /// <summary>
            ///
            /// </summary>
            [NotNull] public const string Target = "Target";
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
            [NotNull] public const string Relationship = "Relationship";
        }
    }
}