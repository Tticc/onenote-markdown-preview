using System;
using System.Collections.Generic;
using System.Xml;
namespace OneNoteMarkdown
{
    public static class ParagraphPolicy
    {
        public static bool IsPlain(XmlElement paragraph,XmlDocument document)
        {
            const string uri="http://schemas.microsoft.com/office/onenote/2013/onenote";
            var ns=new XmlNamespaceManager(document.NameTable);ns.AddNamespace("o",uri);
            if(paragraph.LocalName!="OE"||paragraph.NamespaceURI!=uri)return false;
            if(paragraph.SelectSingleNode("o:List|o:Tag|o:OEChildren",ns)!=null)return false;
            if(!paragraph.HasAttribute("quickStyleIndex"))return true;
            string index=paragraph.GetAttribute("quickStyleIndex");XmlElement definition=null;
            foreach(XmlElement candidate in document.SelectNodes("/o:Page/o:QuickStyleDef",ns))
                if(candidate.GetAttribute("index")==index){if(definition!=null)return false;definition=candidate;}
            if(definition==null)return false;
            // OneNote's exported built-in paragraph style is named p, not a localized UI label.
            if(definition.GetAttribute("name")!="p")return false;
            var attributes=new HashSet<string>{"index","name","fontColor","highlightColor","font","fontSize","spaceBefore","spaceAfter"};
            foreach(XmlAttribute attribute in definition.Attributes)
                if(attribute.NamespaceURI!=""||!attributes.Contains(attribute.Name))return false;
            return true;
        }
    }
}
