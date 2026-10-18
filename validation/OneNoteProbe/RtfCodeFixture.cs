using System;
using System.Globalization;
using System.Text;

// Independent RTF transport experiment, not a Markdown implementation.
internal static class RtfCodeFixture
{
    internal static string Render(string text)
    {
        var output=new StringBuilder(@"{\rtf1\ansi\ansicpg1252\deff0\uc1{\fonttbl{\f0\fmodern Consolas;}}\pard\plain\f0\fs22 ");
        for(int i=0;i<text.Length;i++)
        {
            char value=text[i];
            if(value=='\r')
            {
                if(i+1<text.Length&&text[i+1]=='\n')i++;
                output.Append(@"\line ");
            }
            else if(value=='\n')output.Append(@"\line ");
            else if(value=='\t')output.Append(@"\tab ");
            else if(value=='\\'||value=='{'||value=='}')output.Append('\\').Append(value);
            else if(value>=32&&value<=126)output.Append(value);
            else output.Append(@"\u").Append(unchecked((short)value).ToString(CultureInfo.InvariantCulture)).Append('?');
        }
        // Explicit terminal paragraph is the RTF document terminator, not code text.
        return output.Append(@"\par}").ToString();
    }
}
