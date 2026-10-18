using System;
using System.Collections.Generic;
using System.Xml;

// Categories only: never uses these differences as permission to restore/paste.
internal static class OutsideDiff
{
    internal static string Compare(XmlNode before, XmlNode after)
    {
        var categories = new SortedSet<string>(StringComparer.Ordinal);
        Visit(before, after, categories);
        return categories.Count == 0 ? "none" : String.Join(",", categories);
    }
    private static void Visit(XmlNode a, XmlNode b, ISet<string> result)
    {
        if(a.NodeType!=b.NodeType || a.LocalName!=b.LocalName || a.NamespaceURI!=b.NamespaceURI)
        { result.Add("node-shape");return; }
        if(a is XmlElement)
        {
            var x=(XmlElement)a;var y=(XmlElement)b;
            foreach(XmlAttribute attribute in x.Attributes)
            {
                var other=y.GetAttributeNode(attribute.LocalName,attribute.NamespaceURI);
                if(other==null || other.Value!=attribute.Value) Attribute(x,attribute,result);
            }
            foreach(XmlAttribute attribute in y.Attributes)
                if(x.GetAttributeNode(attribute.LocalName,attribute.NamespaceURI)==null) Attribute(y,attribute,result);
        }
        else if(a.Value!=b.Value) result.Add("content-value");
        if(a.ChildNodes.Count!=b.ChildNodes.Count) {result.Add("child-count");return;}
        for(int i=0;i<a.ChildNodes.Count;i++) Visit(a.ChildNodes[i],b.ChildNodes[i],result);
    }
    private static void Attribute(XmlElement node,XmlAttribute attribute,ISet<string> result)
    {
        string name=attribute.LocalName;
        if(name=="selected")result.Add("selection-metadata");
        else if(name=="lastModifiedTime"||name=="lastModifiedBy"||name=="lastModifiedByInitials")result.Add("modification-metadata");
        else if(node.LocalName=="Position"||node.LocalName=="Size")result.Add("layout-attribute");
        else if(name=="objectID"||name=="ID")result.Add("identity-attribute");
        else if(name=="style")result.Add("style-attribute");
        else result.Add("other-attribute");
    }
}
