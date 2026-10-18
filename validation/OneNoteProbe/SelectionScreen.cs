using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;

// Phase-one structural screening, not permission to paste or a block-boundary classifier.
internal static class SelectionScreen
{
    internal static XmlDocument Parse(string source)
    {
        var settings = new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, MaxCharactersInDocument = 1000000 };
        var document = new XmlDocument { XmlResolver = null };
        using (var input = new StringReader(source))
        using (var reader = XmlReader.Create(input, settings)) document.Load(reader);
        return document;
    }

    internal static string Evaluate(XmlDocument document)
    {
        const string uri = "http://schemas.microsoft.com/office/onenote/2013/onenote";
        if (document.DocumentElement == null || document.DocumentElement.LocalName != "Page" || document.DocumentElement.NamespaceURI != uri)
            return "unsupported-page-schema";
        var ns = new XmlNamespaceManager(document.NameTable);
        ns.AddNamespace("one", uri);
        const string chosen = "[@selected='all' or @selected='partial']";
        if (document.SelectNodes("//one:Title" + chosen, ns).Count != 0) return "title";
        if (document.SelectNodes("//one:Outline" + chosen, ns).Count != 1) return "body-container-count";
        if (document.SelectNodes("//one:Cell" + chosen, ns).Count > 1) return "cross-cell";
        var known = new HashSet<string> { "Page", "Outline", "OEChildren", "OE", "T", "Table", "Row", "Cell" };
        foreach (XmlElement element in document.SelectNodes("//*[@selected]"))
        {
            string state = element.GetAttribute("selected");
            if (state == "none") continue;
            if (state != "all" && state != "partial") return "unknown-selection-state";
            if (element.NamespaceURI != uri || !known.Contains(element.LocalName)) return "selected-object";
            // No single-cell permission established: conservative diagnostic refusal only.
            if (element.LocalName == "Table" || element.LocalName == "Cell") return "table-not-validated";
        }
        int length = 0;
        foreach (XmlElement text in document.SelectNodes("//one:T" + chosen, ns))
        {
            if (text.SelectSingleNode("ancestor::one:Title", ns) != null) return "title";
            XmlElement outline = text.SelectSingleNode("ancestor::one:Outline", ns) as XmlElement;
            if (outline == null || (outline.GetAttribute("selected") != "all" && outline.GetAttribute("selected") != "partial")) return "orphan-text";
            if (text.GetAttribute("selected") != "all") return "partial-text-not-validated";
            XmlDocument fragment;
            try { fragment = FixtureText.Parse(text.InnerText); }
            catch (XmlException) { return "text-decode"; }
            foreach (XmlElement child in fragment.SelectNodes("/root//*"))
                if (child.NamespaceURI != "" || (child.LocalName != "span" && child.LocalName != "br")) return "text-markup-not-validated";
            // Excludes automatic markup such as br from a false empty-selection approval.
            length += fragment.DocumentElement.InnerText.Length;
            XmlElement oe = text.ParentNode as XmlElement;
            if (oe == null || oe.LocalName != "OE" || !oe.HasAttribute("objectID") || String.IsNullOrEmpty(oe.GetAttribute("objectID"))) return "missing-range-identity";
        }
        return length == 0 ? "empty-text" : "candidate-needs-focus-range-style-checks";
    }
}
