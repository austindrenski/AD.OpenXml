using System.Linq;
using System.Xml.Linq;
using AD.IO;
using AD.OpenXml.Properties;
using AD.Xml;
using JetBrains.Annotations;

namespace AD.OpenXml.Documents
{
    [PublicAPI]
    public static class TransferFootnotesExtensions
    {
        private static readonly XNamespace C = XNamespaces.OpenXmlPackageContentTypes;

        private static readonly XNamespace R = XNamespaces.OpenXmlPackageRelationships;

        private static readonly XNamespace S = XNamespaces.OpenXmlOfficeDocumentRelationships;

        private static readonly XNamespace W = XNamespaces.OpenXmlWordprocessingmlMain;

    }
}
