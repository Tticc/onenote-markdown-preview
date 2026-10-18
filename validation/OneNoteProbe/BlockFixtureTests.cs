using System;
using System.Security;
internal static class BlockFixtureTests
{
    private static string Page(string contents){return "<Page xmlns='http://schemas.microsoft.com/office/onenote/2013/onenote'><Outline selected='partial'><OEChildren>"+contents+"</OEChildren></Outline></Page>";}
    private static int Main()
    {
        string body="";int i=0;
        foreach(string line in BlockFixture.Source.Split('\n'))body+="<OE objectID='"+(i++)+"'><T selected='all'>"+SecurityElement.Escape(line)+"</T></OE>";
        if(!BlockFixture.Matches(Page(body)))return 1;
        if(BlockFixture.Matches(Page(body.Replace("测试标题","改变"))))return 1;
        if(BlockFixture.Matches(Page(body.Replace("<T selected='all'>##","<T>前文</T><T selected='all'>##"))))return 1;
        if(BlockFixture.Matches(Page(body.Replace("<T selected='all'>`code`</T>","<T selected='all'>`code`</T><T>后文</T>"))))return 1;
        if(!BlockFixture.Matches(Page(body+"<OE objectID='empty'><T selected='all'></T></OE>")))return 1;
        if(BlockFixture.Matches(Page(body.Replace("<OE objectID='1'>","<OE objectID='gap'><T>gap</T></OE><OE objectID='1'>"))))return 1;
        string fenced="";i=0;
        foreach(string line in BlockFixture.FencedSource.Split('\n'))fenced+="<OE objectID='f"+(i++)+"'><T selected='all'>"+SecurityElement.Escape(line)+"</T></OE>";
        if(!BlockFixture.Matches(Page(fenced),BlockFixture.FencedSource))return 1;
        if(BlockFixture.Matches(Page(fenced.Replace("    alpha","alpha")),BlockFixture.FencedSource))return 1;
        if(BlockFixture.Matches(Page(fenced.Replace("中文😀","中文")),BlockFixture.FencedSource))return 1;
        var markup=FixtureText.Parse(CodeWhitespaceFixture.Render("    alpha\n\n    中文😀"));
        foreach(System.Xml.XmlElement br in markup.SelectNodes("//br"))br.ParentNode.ReplaceChild(markup.CreateTextNode("\n"),br);
        if(markup.DocumentElement.InnerText.Replace('\u00a0',' ')!="    alpha\n\n    中文😀")return 1;
        if(!markup.DocumentElement.InnerText.Contains("\u00a0"))return 1;
        if(FixtureText.Parse(CodeWhitespaceFixture.Render("<&>" )).DocumentElement.InnerText!="<&>")return 1;
        bool rejected=false;try{CodeWhitespaceFixture.Render("\tcode");}catch(NotSupportedException){rejected=true;}if(!rejected)return 1;
        string remaining="";i=0;
        foreach(string line in BlockFixture.RemainingSource.Split('\n'))remaining+="<OE objectID='r"+(i++)+"'><T selected='all'>"+SecurityElement.Escape(line)+"</T></OE>";
        if(!BlockFixture.Matches(Page(remaining),BlockFixture.RemainingSource))return 1;
        if(BlockFixture.Matches(Page(remaining.Replace("[x]","[?]")),BlockFixture.RemainingSource))return 1;
        var rich=FixtureText.Parse(BlockFixture.RemainingHtml);
        if(rich.SelectNodes("//i").Count!=1||rich.SelectNodes("//ol/li").Count!=1||!rich.DocumentElement.InnerText.Contains("☐ 未完成")||!rich.DocumentElement.InnerText.Contains("☑ 已完成"))return 1;
        Console.WriteLine("PASS 16 block/whitespace/style fixture cases; NBSP candidate is visual equivalence only; no input sent");return 0;
    }
}
