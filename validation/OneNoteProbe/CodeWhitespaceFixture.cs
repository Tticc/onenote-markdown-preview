using System;
using System.Text;

// Import candidate only. NBSP is intentional and is NOT U+0020 preservation.
internal static class CodeWhitespaceFixture
{
    internal static string Render(string source)
    {
        if(source.IndexOf('\t')>=0)throw new NotSupportedException("Tab policy requires a separate fidelity decision");
        string[] lines=source.Replace("\r\n","\n").Split('\n');
        var result=new StringBuilder("<div style=\"font-family:Consolas\"><span style=\"font-family:Consolas\">");
        for(int i=0;i<lines.Length;i++)
        {
            if(i>0)result.Append("<br />");
            foreach(char value in lines[i])
            {
                if(value==' ')result.Append("&#160;");
                else if(value=='&')result.Append("&amp;");
                else if(value=='<')result.Append("&lt;");
                else if(value=='>')result.Append("&gt;");
                else result.Append(value);
            }
        }
        return result.Append("</span></div>").ToString();
    }
}
