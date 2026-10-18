using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

// One exact whole-paragraph fixture, not a Markdown implementation.
internal static class BlockFixture
{
    internal const string Source = "## 测试标题\n- 第一项\n> 引用文字\n`code`";
    internal const string Html = "<div><h2>测试标题</h2><ul><li>第一项</li></ul><blockquote>引用文字</blockquote><p><span style=\"font-family:Consolas\">code</span></p></div>";
    internal const string FencedSource = "```\n    alpha\n\n    中文😀\n```";
    internal const string RemainingSource = "*斜体* ~~删除线~~\n1. 有序项\n- [ ] 未完成\n- [x] 已完成";
    internal const string RemainingHtml = "<div><p><i>斜体</i> <span style=\"text-decoration:line-through\">删除线</span></p><ol><li>有序项</li></ol><p>☐ 未完成</p><p>☑ 已完成</p></div>";
    internal const string FencedHtml = "<pre style=\"font-family:Consolas;white-space:pre\">    alpha\n\n    中文😀</pre>";
    internal static string ExplicitWhitespaceHtml { get { return CodeWhitespaceFixture.Render("    alpha\n\n    中文😀"); } }
    internal static bool Matches(string xml)
    { return Matches(xml,Source); }
    internal static bool Matches(string xml,string expected)
    {
        var doc=SelectionScreen.Parse(xml);
        if(SelectionScreen.Evaluate(doc)!="candidate-needs-focus-range-style-checks")return false;
        var ns=new XmlNamespaceManager(doc.NameTable);ns.AddNamespace("o","http://schemas.microsoft.com/office/onenote/2013/onenote");
        var paragraphs=doc.SelectNodes("//o:OE[o:T[@selected='all']]",ns);
        var lines=new List<string>();
        XmlNode parent=null;XmlNode previous=null;
        foreach(XmlElement paragraph in paragraphs)
        {
            if(parent==null)parent=paragraph.ParentNode;
            if(!Object.ReferenceEquals(parent,paragraph.ParentNode))return false;
            if(previous!=null&&!Object.ReferenceEquals(previous.NextSibling,paragraph))return false;
            previous=paragraph;
            var line=new StringBuilder();
            foreach(XmlNode child in paragraph.ChildNodes)
            {
                if(child.NodeType==XmlNodeType.Whitespace)continue;
                var text=child as XmlElement;
                if(text==null||text.LocalName!="T"||text.NamespaceURI!=ns.LookupNamespace("o"))return false;
                string decoded=FixtureText.Parse(text.InnerText).DocumentElement.InnerText;
                // Reject any visible unselected prefix/suffix in the target paragraph.
                if(text.GetAttribute("selected")!="all"&&decoded.Length!=0)return false;
                if(text.GetAttribute("selected")=="all")line.Append(decoded);
            }
            lines.Add(line.ToString());
        }
        int expectedLines=expected.Split('\n').Length;
        if(lines.Count==expectedLines+1&&lines[lines.Count-1].Length==0)lines.RemoveAt(lines.Count-1); // explicit next-line-start endpoint
        return String.Join("\n",lines)==expected;
    }
}
