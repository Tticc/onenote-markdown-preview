using System;
using System.Globalization;
using System.Text;

// Fixed validation fixtures only; not a Markdown parser or product HTML sanitizer.
internal static class HtmlFixturePayload
{
    internal static byte[] Create(string fragment)
    {
        const string prefix = "<html><head><meta charset=\"utf-8\"></head><body><!--StartFragment-->";
        const string suffix = "<!--EndFragment--></body></html>";
        const string template = "Version:1.0\r\nStartHTML:{0:D10}\r\nEndHTML:{1:D10}\r\nStartFragment:{2:D10}\r\nEndFragment:{3:D10}\r\n";
        int header = Encoding.ASCII.GetByteCount(String.Format(CultureInfo.InvariantCulture, template, 0, 0, 0, 0));
        int start = header + Encoding.UTF8.GetByteCount(prefix);
        int end = start + Encoding.UTF8.GetByteCount(fragment);
        int finish = end + Encoding.UTF8.GetByteCount(suffix);
        string text = String.Format(CultureInfo.InvariantCulture, template, header, finish, start, end) + prefix + fragment + suffix;
        byte[] bytes = new UTF8Encoding(false, true).GetBytes(text);
        if (bytes.Length != finish) throw new InvalidOperationException("Offset mismatch");
        return bytes;
    }
}

internal static class HtmlPayloadTests
{
    private static int Offset(byte[] bytes, string name)
    {
        string text = Encoding.ASCII.GetString(bytes, 0, Math.Min(bytes.Length, 160));
        int start = text.IndexOf(name + ":", StringComparison.Ordinal) + name.Length + 1;
        return Int32.Parse(text.Substring(start, 10), CultureInfo.InvariantCulture);
    }
    private static int Main()
    {
        try
        {
            string[] fixtures = {
                "<div><b>测试文字😀</b></div>",
                "<div><h2>标题</h2><ul><li>第一项</li><li>☐ 第二项</li></ul></div>",
                "<div><pre>    缩进\r\n\r\n\tcode &amp; &lt;tag&gt;😀</pre></div>",
                "<div><blockquote>引用</blockquote><span>e\u0301</span></div>"
            };
            foreach (string fragment in fixtures)
            {
                byte[] bytes = HtmlFixturePayload.Create(fragment);
                int a = Offset(bytes, "StartHTML"), b = Offset(bytes, "EndHTML");
                int c = Offset(bytes, "StartFragment"), d = Offset(bytes, "EndFragment");
                if (!(a < c && c <= d && d < b && b == bytes.Length)) throw new InvalidOperationException("Invalid bounds");
                var utf8 = new UTF8Encoding(false, true);
                if (utf8.GetString(bytes, c, d - c) != fragment) throw new InvalidOperationException("Fragment changed");
                if (!utf8.GetString(bytes, a, b - a).StartsWith("<html>")) throw new InvalidOperationException("Missing context");
            }
            Console.WriteLine("PASS 4 CF_HTML fixture byte-offset/Unicode/whitespace roundtrips; no clipboard or OneNote access");
            return 0;
        }
        catch (Exception ex) { Console.WriteLine("FAIL " + ex.Message); return 1; }
    }
}
