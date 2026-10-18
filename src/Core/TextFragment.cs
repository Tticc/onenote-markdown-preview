using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Xml;
namespace OneNoteMarkdown
{
    public static class TextFragment
    {
        public static XmlDocument Parse(string markup)
        {
            if(markup==null||markup.Length>262144)throw new InvalidOperationException("fragment-limit");
            string compatible=Regex.Replace(markup,@"<(?:(?:""[^""]*"")|(?:'[^']*')|[^'""<>])*>",delegate(Match tag)
            {
                return Regex.Replace(tag.Value,@"(""[^""]*""|'[^']*')|(?<name>\s+[A-Za-z_:][A-Za-z0-9_:.-]*)\s*=\s*(?<value>[^\s'""=<>`]+)",delegate(Match attribute)
                {return attribute.Groups["name"].Success?attribute.Groups["name"].Value+"=\""+attribute.Groups["value"].Value+"\"":attribute.Value;},RegexOptions.CultureInvariant,TimeSpan.FromMilliseconds(100));
            },RegexOptions.CultureInvariant,TimeSpan.FromMilliseconds(100));
            var settings=new XmlReaderSettings{DtdProcessing=DtdProcessing.Prohibit,XmlResolver=null,MaxCharactersInDocument=262144};
            var doc=new XmlDocument{XmlResolver=null};
            using(var input=new StringReader("<root>"+compatible.Replace("&nbsp;","&#160;")+"</root>"))using(var reader=XmlReader.Create(input,settings))doc.Load(reader);
            foreach(XmlElement element in doc.SelectNodes("/root//*"))
            {
                if(element.NamespaceURI!=""||element.Name!="span")throw new InvalidOperationException("unsupported-source-markup");
                foreach(XmlAttribute attribute in element.Attributes)
                {
                    if(attribute.Name!="style"&&attribute.Name!="lang")throw new InvalidOperationException("unsupported-source-attribute");
                    if(attribute.Name=="style")
                    {
                        string style=attribute.Value.ToLowerInvariant();
                        if(style.Contains("url(")||style.Contains("expression")||style.Contains("font-weight:bold")||style.Contains("font-style:italic")||style.Contains("text-decoration"))throw new InvalidOperationException("source-rich-format-not-supported");
                    }
                }
            }
            return doc;
        }
        public static string Read(string markup){return Parse(markup).DocumentElement.InnerText;}
        public static bool HasContent(string markup)
        {
            if(string.IsNullOrEmpty(markup))return false;
            string text=Regex.Replace(markup,@"<(?:(?:""[^""]*"")|(?:'[^']*')|[^'""<>])*>","",RegexOptions.CultureInvariant,TimeSpan.FromMilliseconds(100));
            return text.Trim().Length!=0;
        }
    }
}
