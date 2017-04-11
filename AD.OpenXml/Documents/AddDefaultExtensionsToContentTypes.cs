using System.Linq;
using System.Xml.Linq;
using AD.IO;
using AD.Xml;
using JetBrains.Annotations;

namespace AD.OpenXml.Documents
{
    [PublicAPI]
    public static class AddDefaultExtensionsToContentTypesExtensions
    {
        private static readonly XNamespace C = XNamespaces.OpenXmlPackageContentTypes;

        public static void AddDefaultExtensionsToContentTypes(this DocxFilePath toFilePath)
        {
            XElement types = toFilePath.ReadAsXml("[Content_Types].xml");

            if (types.Descendants().Attributes("Extension").All(x => x.Value != "docx"))
            {
                XElement docx =
                    new XElement(C + "Default",
                        new XAttribute("Extension", "docx"),
                        new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"));
                types.AddFirst(docx);
            }

            if (types.Descendants().Attributes("Extension").All(x => x.Value != "xlsx"))
            {
                XElement xlsx =
                    new XElement(C + "Default",
                        new XAttribute("Extension", "xlsx"),
                        new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));
                types.AddFirst(xlsx);
            }

            types.WriteInto(toFilePath, "[Content_Types].xml");
        }
    }
}
