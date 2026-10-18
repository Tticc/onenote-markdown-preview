using System;
using System.Xml;

internal static class SelectionScreenTests
{
    private static string Page(string content)
    { return "<Page xmlns='http://schemas.microsoft.com/office/onenote/2013/onenote'>" + content + "</Page>"; }
    private static string Outline(string content)
    { return "<Outline selected='partial'><OEChildren>" + content + "</OEChildren></Outline>"; }
    private static string Text(string content)
    { return "<OE objectID='fixture' selected='partial'><T selected='all'><![CDATA[" + content + "]]></T></OE>"; }
    private static int count;
    private static void Check(string source, string expected)
    {
        string actual = SelectionScreen.Evaluate(SelectionScreen.Parse(source));
        if (actual != expected) throw new InvalidOperationException("Expected " + expected + "; got " + actual);
        count++;
    }
    private static int Main()
    {
        try
        {
            Check(Page(Outline(Text("**text**"))), "candidate-needs-focus-range-style-checks");
            Check(Page(Outline(Text("<span style='font-family:Calibri'>中文😀</span>"))), "candidate-needs-focus-range-style-checks");
            Check(Page(Outline(Text("<span lang=zh-CN style='font-size:11pt'>**text**</span>"))), "candidate-needs-focus-range-style-checks");
            Check(Page(Outline(Text("<span style='x=y'>a=b</span>"))), "candidate-needs-focus-range-style-checks");
            if (FixtureText.Parse("<span lang=zh-CN>a=b &amp; c</span>").DocumentElement.InnerText != "a=b & c") throw new InvalidOperationException("Fixture text normalization changed text");
            count++;
            Check(Page(Outline(Text(""))), "empty-text");
            Check(Page(Outline(Text("")) + Outline(Text(""))), "body-container-count");
            Check(Page("<Title selected='partial'>" + Text("title") + "</Title>"), "title");
            Check(Page(Outline(Text("x") + "<OE><Image selected='all'/></OE>")), "selected-object");
            Check(Page(Outline(Text("x") + "<OE><Image/></OE>")), "candidate-needs-focus-range-style-checks");
            Check(Page(Outline("<Table selected='partial'><Cell selected='all'/><Cell selected='all'/></Table>")), "cross-cell");
            Check(Page(Outline("<Table selected='partial'><Cell selected='all'/></Table>")), "table-not-validated");
            Check(Page(Outline(Text("<a href='x'>link</a>"))), "text-markup-not-validated");
            Check(Page(Outline(Text("<unknown>text</unknown>"))), "text-markup-not-validated");
            Check(Page(Outline(Text("<span>broken"))), "text-decode");
            Check(Page(Outline(Text("x"))).Replace("objectID='fixture'", ""), "missing-range-identity");
            Check(Page(Outline(Text("x"))).Replace("selected='all'", "selected='unexpected'"), "unknown-selection-state");
            Check(Page(Outline(Text("x"))).Replace("<T selected='all'", "<T selected='partial'"), "partial-text-not-validated");
            Check("<Page/>", "unsupported-page-schema");
            try { SelectionScreen.Parse("<!DOCTYPE root [<!ENTITY x SYSTEM 'file:///not-read'>]><root>&x;</root>"); throw new InvalidOperationException("DTD accepted"); }
            catch (XmlException) { count++; }
            try { SelectionScreen.Parse("<root>" + new string('x', 1000001) + "</root>"); throw new InvalidOperationException("Oversize accepted"); }
            catch (XmlException) { count++; }
            string valid = Page(Outline(Text("**text**")));
            var evidence = SelectionEvidence.Capture(valid, "page-a", 10, 11, 10, "OneNote::DocumentCanvas", 1);
            if (evidence == null || !evidence.Matches(SelectionEvidence.Capture(valid, "page-a", 10, 11, 10, "OneNote::DocumentCanvas", 1))) throw new InvalidOperationException("Identical evidence mismatch");
            count++;
            if (evidence.Matches(SelectionEvidence.Capture(valid, "page-b", 10, 11, 10, "OneNote::DocumentCanvas", 1))) throw new InvalidOperationException("Page switch accepted");
            count++;
            if (evidence.Matches(SelectionEvidence.Capture(valid.Replace("**text**", "**edited**"), "page-a", 10, 11, 10, "OneNote::DocumentCanvas", 1))) throw new InvalidOperationException("Content change accepted");
            count++;
            if (SelectionEvidence.Capture(valid, "page-a", 10, 11, 10, "RICHEDIT60W", 1) != null) throw new InvalidOperationException("Search focus accepted");
            count++;
            if (SelectionEvidence.Capture(valid, "page-a", 10, 11, 99, "OneNote::DocumentCanvas", 1) != null) throw new InvalidOperationException("Other foreground accepted");
            count++;
            if (SelectionEvidence.Capture(valid, "page-a", 10, 0, 10, "OneNote::DocumentCanvas", 1) != null) throw new InvalidOperationException("Missing focus accepted");
            count++;
            if (SelectionEvidence.Capture(valid, "page-a", 10, 11, 10, "OneNote::DocumentCanvas", 2) != null) throw new InvalidOperationException("Multiple windows accepted");
            count++;
            string selectedFirst = Page(Outline("<OE objectID='fixture'><T selected='all'>same</T><T> gap same</T></OE>")).Replace("<OE objectID", "<OE selected='partial' objectID");
            string selectedSecond = Page(Outline("<OE objectID='fixture'><T>same gap </T><T selected='all'>same</T></OE>")).Replace("<OE objectID", "<OE selected='partial' objectID");
            var firstEvidence = SelectionEvidence.Capture(selectedFirst, "page-a", 10, 11, 10, "OneNote::DocumentCanvas", 1);
            var secondEvidence = SelectionEvidence.Capture(selectedSecond, "page-a", 10, 11, 10, "OneNote::DocumentCanvas", 1);
            if (firstEvidence == null || secondEvidence == null || firstEvidence.Matches(secondEvidence)) throw new InvalidOperationException("Repeated-text move accepted");
            count++;
            Console.WriteLine("PASS " + count + " structural screening cases; no paste authorization");
            return 0;
        }
        catch (Exception ex) { Console.WriteLine("FAIL " + ex.Message); return 1; }
    }
}
