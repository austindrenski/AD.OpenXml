using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using AD.Xml;
using JetBrains.Annotations;

namespace AD.OpenXml
{
    /// <summary>
    ///
    /// </summary>
    [PublicAPI]
    public class ContentTypes
    {
        [NotNull] private static readonly XNamespace T = XNamespaces.OpenXmlPackageContentTypes;

        /// <summary>
        ///
        /// </summary>
        [NotNull]
        public IEnumerable<Default> Defaults { get; }

        /// <summary>
        ///
        /// </summary>
        [NotNull]
        public IEnumerable<Override> Overrides { get; }

        /// <summary>
        ///
        /// </summary>
        /// <returns></returns>
        [Pure]
        [NotNull]
        public XElement ToXElement()
        {
            return
                new XElement(
                    T + "Types",
                    Defaults.Select(x => x.ToXElement()),
                    Overrides.Select(x => x.ToXElement()));
        }

        /// <summary>
        ///
        /// </summary>
        [PublicAPI]
        public readonly struct Default
        {
            /// <summary>
            ///
            /// </summary>
            /// <returns></returns>
            [Pure]
            [NotNull]
            public XElement ToXElement()
            {
                return
                    new XElement(
                        T + "Default");
            }
        }

        /// <summary>
        ///
        /// </summary>
        [PublicAPI]
        public readonly struct Override
        {
            /// <summary>
            ///
            /// </summary>
            /// <returns></returns>
            [Pure]
            [NotNull]
            public XElement ToXElement()
            {
                return
                    new XElement(
                        T + "Override");
            }
        }
    }
}