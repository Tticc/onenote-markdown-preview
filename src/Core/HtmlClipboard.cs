using System;
using System.Globalization;
using System.Text;
namespace OneNoteMarkdown
{
    public static class HtmlClipboard
    {
        public static byte[] Create(string fragment)
        {
            MarkdownConverter.ValidateHtml(fragment);
            const string prefix="<html><head><meta charset=\"utf-8\"></head><body><!--StartFragment-->";
            const string suffix="<!--EndFragment--></body></html>";
            const string header="Version:1.0\r\nStartHTML:{0:D10}\r\nEndHTML:{1:D10}\r\nStartFragment:{2:D10}\r\nEndFragment:{3:D10}\r\n";
            int start=String.Format(CultureInfo.InvariantCulture,header,0,0,0,0).Length;
            int first=start+Encoding.UTF8.GetByteCount(prefix),last=first+Encoding.UTF8.GetByteCount(fragment),end=last+Encoding.UTF8.GetByteCount(suffix);
            return new UTF8Encoding(false,true).GetBytes(String.Format(CultureInfo.InvariantCulture,header,start,end,first,last)+prefix+fragment+suffix+"\0");
        }
    }
}
