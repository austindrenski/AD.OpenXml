using JetBrains.Annotations;

namespace AD.OpenXml.Structure
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
        [NotNull] public const string Namespace = "http://schemas.openxmlformats.org/package/2006/content-types";

        /// <summary>
        ///
        /// </summary>
        [NotNull] public const string Root = "Types";

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
            [NotNull] public const string ContentType = "ContentType";

            /// <summary>
            ///
            /// </summary>
            [NotNull] public const string Extension = "Extension";

            /// <summary>
            ///
            /// </summary>
            [NotNull] public const string PartName = "PartName";
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
            [NotNull] public const string Default = "Default";

            /// <summary>
            ///
            /// </summary>
            [NotNull] public const string Override = "Override";
        }
    }
}