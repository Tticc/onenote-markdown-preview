using System;
using System.Text.RegularExpressions;
using System.Xml;

// Narrow diagnostic adapter for OneNote span wrappers, not a general HTML parser.
internal static class FixtureText
{
    internal static XmlDocument Parse(string markup)
    {
        if (markup.Length > 65536) throw new XmlException("Fragment limit");
        // Quoted attribute values are consumed unchanged, so text resembling an
        // assignment inside a quoted style is never rewritten. Only tag lexemes
        // are visited; text nodes are never transformed.
        string compatible = Regex.Replace(markup, @"<(?:(?:""[^""]*"")|(?:'[^']*')|[^'""<>])*>", delegate(Match tag)
        {
            return Regex.Replace(tag.Value, @"(""[^""]*""|'[^']*')|(?<name>\s+[A-Za-z_:][A-Za-z0-9_:.-]*)\s*=\s*(?<value>[^\s'""=<>`]+)", delegate(Match attribute)
            {
                if (!attribute.Groups["name"].Success) return attribute.Value;
                return attribute.Groups["name"].Value + "=\"" + attribute.Groups["value"].Value + "\"";
            }, RegexOptions.CultureInvariant, TimeSpan.FromMilliseconds(100));
        }, RegexOptions.CultureInvariant, TimeSpan.FromMilliseconds(100));
        return SelectionScreen.Parse("<root>" + compatible.Replace("&nbsp;", "&#160;") + "</root>");
    }
}
