using System;
using System.Text;
using System.Xml;

// Fixed inline fixture observer, not a general paste-completion transaction.
internal static class InlinePostcondition
{
    private const string Uri = "http://schemas.microsoft.com/office/onenote/2013/onenote";
    internal static bool BaselineMatches(string selectedXml,string basicXml)
    {
        var selected=SelectionScreen.Parse(selectedXml);var basic=SelectionScreen.Parse(basicXml);
        var ns=new XmlNamespaceManager(selected.NameTable);ns.AddNamespace("o",Uri);
        var nodes=selected.SelectNodes("//o:T[@selected='all']",ns);
        if(nodes.Count!=1)return false;
        var target=nodes[0].ParentNode as XmlElement;
        if(target==null||!target.HasAttribute("objectID"))return false;
        var otherNs=new XmlNamespaceManager(basic.NameTable);otherNs.AddNamespace("o",Uri);
        XmlElement found=null;
        foreach(XmlElement candidate in basic.SelectNodes("//o:OE[@objectID]",otherNs))
            if(candidate.GetAttribute("objectID")==target.GetAttribute("objectID")){if(found!=null)return false;found=candidate;}
        if(found==null)return false;
        var a=new StringBuilder();var b=new StringBuilder();
        foreach(XmlElement text in target.ChildNodes){if(text.LocalName!="T"||text.NamespaceURI!=Uri)return false;a.Append(FixtureText.Parse(text.InnerText).DocumentElement.InnerText);}
        foreach(XmlElement text in found.ChildNodes){if(text.LocalName!="T"||text.NamespaceURI!=Uri)return false;b.Append(FixtureText.Parse(text.InnerText).DocumentElement.InnerText);}
        return a.ToString()==b.ToString();
    }
    internal static bool Matches(string before, string after)
    { return Diagnose(before, after) == "expected-text-and-outside-serialization-match"; }

    internal static string Diagnose(string before, string after)
    { return Diagnose(before,after,null); }

    internal static string Diagnose(string before, string after,string basicBefore)
    {
        var a = SelectionScreen.Parse(before); var b = SelectionScreen.Parse(after);
        var ns = new XmlNamespaceManager(a.NameTable); ns.AddNamespace("o", Uri);
        var selected = a.SelectNodes("//o:T[@selected='all']", ns);
        if (selected.Count != 1) return "original-selection-count";
        var target = selected[0].ParentNode as XmlElement;
        if (target == null || !target.HasAttribute("objectID")) return "original-target-identity";
        string id = target.GetAttribute("objectID");
        var beforeText = new StringBuilder(); var expected = new StringBuilder();
        foreach (XmlElement node in target.ChildNodes)
        {
            if (node.LocalName != "T" || node.NamespaceURI != Uri) return "original-target-nontext-child";
            string text = FixtureText.Parse(node.InnerText).DocumentElement.InnerText;
            beforeText.Append(text);
            if (Object.ReferenceEquals(node, selected[0]))
            {
                if (text != "**测试文字**") return "original-token-mismatch";
                expected.Append("测试文字");
            }
            else expected.Append(text);
        }
        var afterNs = new XmlNamespaceManager(b.NameTable); afterNs.AddNamespace("o", Uri);
        XmlElement found = null;
        foreach (XmlElement candidate in b.SelectNodes("//o:OE[@objectID]", afterNs))
            if (candidate.GetAttribute("objectID") == id) { if (found != null) return "duplicate-target-identity"; found = candidate; }
        if (found == null) return "target-identity-not-found";
        var actual = new StringBuilder();
        foreach (XmlElement node in found.ChildNodes)
        {
            if (node.LocalName != "T" || node.NamespaceURI != Uri) return "result-target-nontext-child";
            actual.Append(FixtureText.Parse(node.InnerText).DocumentElement.InnerText);
        }
        if (actual.ToString() == beforeText.ToString()) return "target-text-unchanged";
        if (actual.ToString() != expected.ToString()) return "target-text-unexpected";
        // Compare all content outside the target after stripping selection-only metadata.
        if(basicBefore!=null)
        {
            a=SelectionScreen.Parse(basicBefore);
            var baselineNs=new XmlNamespaceManager(a.NameTable);baselineNs.AddNamespace("o",Uri);
            XmlElement baselineTarget=null;
            foreach(XmlElement candidate in a.SelectNodes("//o:OE[@objectID]",baselineNs))
                if(candidate.GetAttribute("objectID")==id){if(baselineTarget!=null)return "duplicate-baseline-target";baselineTarget=candidate;}
            if(baselineTarget==null)return "missing-baseline-target";
            var baselineText=new StringBuilder();
            foreach(XmlElement node in baselineTarget.ChildNodes)
            {
                if(node.LocalName!="T"||node.NamespaceURI!=Uri)return "baseline-target-nontext-child";
                baselineText.Append(FixtureText.Parse(node.InnerText).DocumentElement.InnerText);
            }
            if(baselineText.ToString()!=beforeText.ToString())return "baseline-target-text-changed";
            target=baselineTarget;
        }
        target.ParentNode.RemoveChild(target); found.ParentNode.RemoveChild(found);
        Normalize(a); Normalize(b);
        return a.OuterXml == b.OuterXml ? "expected-text-and-outside-serialization-match" : "expected-target-text-but-outside-serialization-differs:" + OutsideDiff.Compare(a,b);
    }
    private static void Normalize(XmlDocument doc)
    {
        foreach (XmlElement element in doc.SelectNodes("//*"))
        {
            element.RemoveAttribute("selected");
            element.RemoveAttribute("lastModifiedTime");
            element.RemoveAttribute("lastModifiedBy");
            element.RemoveAttribute("lastModifiedByInitials");
        }
    }
}
