using System.Xml.Linq;
using AD.Xml;
using JetBrains.Annotations;

namespace AD.OpenXml.Elements
{
    /// <summary>
    /// 
    /// </summary>
    [PublicAPI]
    public static class AddTableCaptionExtensions
    {
        private static readonly XNamespace W = XNamespaces.OpenXmlWordprocessingmlMain;

        private static readonly XNamespace Xml = XNamespace.Xml;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="element"></param>
        public static void AddTableCaption(this XElement element)
        {
            string style = element.Value.Contains("[APPENDIX]") ? "\"Heading 9\"" : "1";

            XElement runProperies =
                new XElement(W + "rPr",
                    new XElement(W + "rStyle",
                        new XAttribute(W + "val", "Strong")));
            XElement fieldCharBegin =
                new XElement(W + "fldChar",
                    new XAttribute(W + "fldCharType", "begin"));
            XElement fieldCharSeparate =
                new XElement(W + "fldChar",
                    new XAttribute(W + "fldCharType", "separate"));
            XElement fieldCharEnd =
                new XElement(W + "fldChar",
                    new XAttribute(W + "fldCharType", "end"));
            XAttribute preserve = new XAttribute(Xml + "space", "preserve");

            XElement label0 =
                new XElement(W + "r",
                    runProperies,
                    new XElement(W + "t", preserve, "Table "));
            XElement label1 =
                new XElement(W + "r",
                    runProperies,
                    fieldCharBegin);
            XElement label2 =
                new XElement(W + "r",
                    runProperies,
                    new XElement(W + "instrText", preserve, $" STYLEREF {style} \\s "));
            XElement label3 =
                new XElement(W + "r",
                    runProperies,
                    fieldCharSeparate);
            XElement label4 =
                new XElement(W + "r",
                    runProperies,
                    new XElement(W + "t", "0"));
            XElement label5 =
                new XElement(W + "r",
                    runProperies,
                    fieldCharEnd);
            XElement label6 =
                new XElement(W + "r",
                    runProperies,
                    new XElement(W + "t", "."));
            XElement label7 =
                new XElement(W + "r",
                    runProperies,
                    fieldCharBegin);
            XElement label8 =
                new XElement(W + "r",
                    runProperies,
                    new XElement(W + "instrText", preserve, $" SEQ Table \\* ARABIC \\s {style} "));
            XElement label9 =
                new XElement(W + "r",
                    runProperies,
                    fieldCharSeparate);
            XElement label10 =
                new XElement(W + "r",
                    runProperies,
                    new XElement(W + "t", "0"));
            XElement label11 =
                new XElement(W + "r",
                    runProperies,
                    fieldCharEnd);
            XElement label12 =
                new XElement(W + "r",
                    runProperies,
                    new XElement(W + "t", preserve, " "));
            element.AddFirst(
                label0,
                label1,
                label2,
                label3,
                label4,
                label5,
                label6,
                label7,
                label8,
                label9,
                label10,
                label11,
                label12);

            foreach (XElement text in element.Descendants(W + "t"))
            {
                text.Value = text.Value.Replace("[", null);
                text.Value = text.Value.Replace("APPENDIX", null);
                text.Value = text.Value.Replace("]", null);
            }
        }
    }
}