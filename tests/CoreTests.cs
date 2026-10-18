using System;
using System.Text;
using OneNoteMarkdown;
internal static class CoreTests
{
    private static int count;
    private static void Check(bool pass,string label){if(!pass)throw new Exception(label);count++;}
    private static ConversionResult Convert(string text,bool blocks=true){return new MarkdownConverter().Convert(text,blocks);}
    private static int Main()
    {
        try
        {
            var bold=Convert("前文 **测试文字** 后文",false);Check(bold.CanApply&&bold.Html.Contains("<b>测试文字</b>")&&bold.Text=="前文 测试文字 后文","inline bold fidelity");
            Check(Convert("*斜体* ~~删除线~~",false).CanApply,"italic strike");
            var code=Convert("前文 `code` 后文",false);Check(code.CanApply&&code.Html.Contains("background-color:#eeeeee"),"code visual style");
            var heading=Convert("# 标题\n\n- 第一项\n- 第二项");Check(heading.CanApply,"heading list: "+heading.Status);
            Check(Convert("1. first\n2. second").CanApply,"ordered list");
            var task=Convert("- [ ] 未完成\n- [x] 已完成");Check(task.CanApply&&task.Text.Contains("☐")&&task.Text.Contains("☑"),"task symbols");
            Check(Convert("> 引用").CanApply,"quote");
            Check(Convert("普通文字").Status=="no-markdown","no-op");
            Check(!Convert("**正常** ![图](x)").CanApply,"mixed unsupported rejects whole");
            Check(!Convert("**正常** [链接](https://example.com)").CanApply,"link rejected");
            Check(!Convert("<script>x</script>").CanApply,"raw HTML rejected");
            Check(!Convert("|a|b|\n|-|-|\n|x|y|").CanApply,"table rejected");
            Check(!Convert("```\n    code\n```").CanApply,"fenced code refused");
            Check(!Convert("# 标题",false).CanApply,"partial block refused");
            Check(!Convert("#### 标题").CanApply,"unsupported heading");
            Check(!Convert("**closed** and **unclosed").CanApply,"partial parse not applied");
            Check(!Convert("**closed** and *unclosed").CanApply,"unclosed single delimiter");
            Check(Convert(@"\*literal\* **bold**",false).CanApply,"escaped stars literal");
            Check(!Convert("- outer\n  - inner").CanApply,"nested list refused");
            Check(!Convert("__unsupported__").CanApply,"underscore emphasis refused");
            Check(!Convert("3. item").CanApply,"non-default ordered start rejected");
            Check(!Convert("* 项目").CanApply,"star bullet refused");
            Check(!Convert("+ 项目").CanApply,"plus bullet refused");
            Check(!Convert("标题\n====").CanApply,"setext heading refused");
            Check(!Convert("#### 四级").CanApply,"level-four heading refused");
            Check(Convert("# 标题 **粗体**").CanApply,"heading with inline bold");
            Check(Convert("> 引用 **粗体**\n> 第二行").CanApply,"multiline quote with inline");
            Check(!Convert("普通文字\0尾").CanApply,"NUL character rejected");
            Check(Convert("```\ncode\n```").Status=="fenced-code-fidelity-unverified","fence refusal reason recorded");
            Check(!Convert("无闭合 ```\ncode").CanApply,"unclosed fence refused");
            Check(!Convert("~~~\ncode\n~~~").CanApply,"tilde fence refused");
            Check(!Convert("    code\n\n- 项目").CanApply,"indented code with list refused");
            Check(!Convert("- 项目\n\n```\ncode\n```\n\n**粗体**").CanApply,"fenced code mixed with syntax rejects whole");
            Check(!Convert("- 项目\n\n- 项目二",false).CanApply,"partial block selection needs whole paragraphs");
            Check(Convert("- 项目\n\n- 项目二",false).Status=="whole-paragraphs-required","partial block reason recorded");
            Check(!Convert("**粗体** `代码` **未闭合").CanApply,"dangling marker across inline types refused");
            Check(Convert("**粗体 与 `代码` 混合**",false).CanApply,"bold wrapping inline code");
            Check(Convert("`<b> & *literal*`",false).Text=="<b> & *literal*","code content not reinterpreted");
            Check(!Convert("普通 **粗体** 尾\ud83d\ude00 **未闭合").CanApply,"dangling marker with emoji refused");            string paragraphXml="<Page xmlns='http://schemas.microsoft.com/office/onenote/2013/onenote'><QuickStyleDef index='0' name='p' font='Calibri' fontSize='11.0' fontColor='automatic' highlightColor='automatic' spaceBefore='0.0' spaceAfter='0.0'/><Outline><OEChildren><OE quickStyleIndex='0'><T>text</T></OE></OEChildren></Outline></Page>";
            var paragraphDocument=new System.Xml.XmlDocument{XmlResolver=null};paragraphDocument.LoadXml(paragraphXml);
            var paragraph=(System.Xml.XmlElement)paragraphDocument.GetElementsByTagName("OE")[0];
            Check(ParagraphPolicy.IsPlain(paragraph,paragraphDocument),"default p QuickStyle is plain");
            ((System.Xml.XmlElement)paragraphDocument.GetElementsByTagName("QuickStyleDef")[0]).SetAttribute("name","h1");
            Check(!ParagraphPolicy.IsPlain(paragraph,paragraphDocument),"heading QuickStyle refused");
            paragraphDocument.LoadXml(paragraphXml.Replace("name='p'","name='p' isBold='true'"));paragraph=(System.Xml.XmlElement)paragraphDocument.GetElementsByTagName("OE")[0];
            Check(!ParagraphPolicy.IsPlain(paragraph,paragraphDocument),"unknown style semantics refused");
            paragraphDocument.LoadXml(paragraphXml.Replace("<T>text</T>","<List/><T>text</T>"));paragraph=(System.Xml.XmlElement)paragraphDocument.GetElementsByTagName("OE")[0];
            Check(!ParagraphPolicy.IsPlain(paragraph,paragraphDocument),"native List remains refused");
            paragraphDocument.LoadXml(paragraphXml);paragraph=(System.Xml.XmlElement)paragraphDocument.GetElementsByTagName("OE")[0];paragraph.SetAttribute("quickStyleIndex","99");
            Check(!ParagraphPolicy.IsPlain(paragraph,paragraphDocument),"missing style refused");
            Check(!Convert("1) item").CanApply,"non-dot numbering rejected");
            Check(!Convert("normal &copy; **bold**").CanApply,"unsupported entity not silently changed");
            Check(Convert("`<tag> & **literal**`",false).Text=="<tag> & **literal**","code encoding");
            Check(!Convert("    code\n\n**bold**").CanApply,"indented code whole selection refused");
            var wrapped=TextFragment.Read("<span lang=zh-CN style='font-family:Calibri'>测试😀</span>");Check(wrapped=="测试😀","OneNote source wrapper");
            bool reject=false;try{TextFragment.Read("<a href='x'>text</a>");}catch(Exception){reject=true;}Check(reject,"source hyperlink refused");
            reject=false;try{TextFragment.Read("<!DOCTYPE x [<!ENTITY x SYSTEM 'file:///no-read'>]>&x;");}catch(Exception){reject=true;}Check(reject,"source DTD refused");
            reject=false;try{TextFragment.Read("<span style='font-weight:bold'>x</span>");}catch(Exception){reject=true;}Check(reject,"styled selected run refused");
            Check(TextFragment.HasContent("<span style='font-weight:bold'>x</span>"),"styled sibling run counts as content");
            Check(!TextFragment.HasContent("<span style='font-family:Calibri'></span>")&&!TextFragment.HasContent("  ")&&!TextFragment.HasContent(""),"empty sibling content detected");
            Check(Convert("`**literal**`",false).CanApply,"code syntax literal");
            Check(Convert("**中文😀 e\u0301**",false).Text=="中文😀 e\u0301","Unicode fidelity");
            Check(!Convert(new string('a',32769)).CanApply,"size limit");
            byte[] payload=HtmlClipboard.Create(bold.Html);Check(Encoding.UTF8.GetString(payload).Contains("<!--StartFragment-->"),"HTML envelope");
            reject=false;try{HtmlClipboard.Create("<img src='x'/>");}catch(Exception){reject=true;}Check(reject,"output validation");
            var state=new OperationState();Check(state.Begin()==0,"default disabled");Check(state.Enable(),"enable");long ticket=state.Begin();Check(ticket!=0&&state.Begin()==0,"single flight");state.Disable();Check(!state.Valid(ticket)&&!state.Enable(),"cancel and no overlapping restart");state.End();Check(state.Enable(),"recover after done");state.Fault();Check(!state.Enabled,"fault stops");
            var flow=new OperationState();Check(flow.Enable(),"fault-injection enable");long t1=flow.Begin();Check(t1!=0,"op begin");Check(flow.Begin()==0,"duplicate request rejected");flow.Disable();Check(!flow.Valid(t1),"late response after disable rejected");Check(flow.Begin()==0,"no new op while busy");flow.End();Check(flow.Enable(),"manual re-enable after cancel");long t3=flow.Begin();Check(t3>t1&&flow.Valid(t3),"generation advanced after restart");flow.Fault();Check(!flow.Enabled&&flow.Begin()==0&&!flow.Valid(t3),"fault stops and invalidates in-flight");flow.End();Check(flow.Enable(),"manual recovery after fault");Check(flow.Begin()!=0,"op allowed after recovery");Check(OperationState.Delay(500)==500&&OperationState.Delay(200)==200&&OperationState.Delay(300)==200&&OperationState.Delay(0)==200&&OperationState.Delay(-1)==200,"restore delay policy 200/500 only");
            Console.WriteLine("PASS "+count+" actual core conversion/state tests; no Office/clipboard access");return 0;
        }
        catch(Exception error){Console.WriteLine("FAIL "+error.Message);return 1;}
    }
}
