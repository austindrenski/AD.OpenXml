﻿using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using JetBrains.Annotations;

namespace AjdExtensions.Html
{
    [PublicAPI]
    public static class ConvertStyleToClassExtensions
    {
        public static XElement ConvertStyleToClass(this XElement element)
        {
            IEnumerable<XElement> items =
                element.Descendants()
                       .Where(x => x.Name == "pStyle" || x.Name == "rStyle")
                       .Where(x => x.Attributes("val").All(y => y.Value != "Strong" && y.Value != "Emphasis"))
                       .ToArray();

            foreach (XElement item in items)
            {
                item.Parent?.Parent?.SetAttributeValue("class", item.Attribute("val")?.Value);
                item.Remove();
            }
            return element;
        }
    }
}