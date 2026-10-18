using System;
using System.Windows.Forms;

internal static class RtfCodeFixtureTests
{
    [STAThread]private static int Main()
    {
        string[] cases={"    alpha\n\n    中文😀","\talpha\n\t中文😀","  a  b  ","{code} \\ test","\n\n","e\u0301\n\u00a0","\r\nalpha\r\n","","😀"};
        try
        {
            using(var box=new RichTextBox())
            for(int index=0;index<cases.Length;index++)
            {
                string value=cases[index];
                string rtf=RtfCodeFixture.Render(value);
                foreach(char c in rtf)if(c>127)throw new InvalidOperationException("RTF not ASCII-safe");
                box.Rtf=rtf;
                string expected=value.Replace("\r\n","\n").Replace("\r","\n");
                if(box.Text!=expected)throw new InvalidOperationException("RTF fixture text mismatch case="+index+" actual="+String.Join(",",Array.ConvertAll(box.Text.ToCharArray(),c=>((int)c).ToString("X4")))+" expected="+String.Join(",",Array.ConvertAll(expected.ToCharArray(),c=>((int)c).ToString("X4"))));
            }
            Console.WriteLine("PASS 9 RTF code whitespace/Unicode roundtrips through Windows RichEdit; no OneNote or clipboard access");return 0;
        }
        catch(Exception ex){Console.WriteLine("FAIL "+ex.GetType().Name+" "+ex.Message);return 1;}
    }
}
