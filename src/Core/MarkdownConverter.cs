using System;
using System.Collections.Generic;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using Markdig;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;

namespace OneNoteMarkdown
{
    public sealed class ConversionResult
    {
        public string Status { get; internal set; }
        public string Html { get; internal set; }
        public string Text { get; internal set; }
        public bool HasBlocks { get; internal set; }
        public bool CanApply { get { return Status=="converted"; } }
    }
    public sealed class MarkdownConverter
    {
        public const int MaxCharacters=32768, MaxDepth=16, MaxNodes=4096, MaxOutputBytes=262144;
        private readonly MarkdownPipeline pipeline=new MarkdownPipelineBuilder().UseEmphasisExtras().UseTaskLists().UsePipeTables().Build();
        private int nodes,formats;
        private readonly StringBuilder html=new StringBuilder(),text=new StringBuilder();
        private string source;
        private sealed class Refused:Exception {internal Refused(string category):base(category){}}
        // One converter instance per operation: mutable renderer state is never shared.
        public ConversionResult Convert(string markdown,bool wholeParagraphs)
        { lock(this) { return ConvertLocked(markdown,wholeParagraphs); } }
        private ConversionResult ConvertLocked(string markdown,bool wholeParagraphs)
        {
            if(markdown==null)return Reject("invalid-input");
            if(markdown.Length>MaxCharacters)return Reject("input-limit");
            if(markdown.IndexOf('\0')>=0)return Reject("invalid-character");
            html.Clear();text.Clear();nodes=0;formats=0;source=markdown.Replace("\r\n","\n").Replace('\r','\n');
            try
            {
                // Conservative recognition of unimplemented block constructs.
                if(Regex.IsMatch(source,@"(?m)^ {0,3}(`{3,}|~{3,})",RegexOptions.CultureInvariant,TimeSpan.FromMilliseconds(100)))throw new Refused("fenced-code-fidelity-unverified");
                var document=Markdown.Parse(source,pipeline);
                bool blocks=document.Count>1;
                foreach(var block in document)if(!(block is ParagraphBlock))blocks=true;
                if(blocks&&!wholeParagraphs)throw new Refused("whole-paragraphs-required");
                if(!wholeParagraphs&&document.Count==1&&document[0] is ParagraphBlock)
                    Inline(((ParagraphBlock)document[0]).Inline,0);
                else {html.Append("<div>");foreach(var block in document)Block(block,0);html.Append("</div>");}
                if(formats==0)return Reject("no-markdown");
                string fragment=html.ToString();
                ValidateHtml(fragment);
                if(Encoding.UTF8.GetByteCount(fragment)>MaxOutputBytes)throw new Refused("output-limit");
                return new ConversionResult{Status="converted",Html=fragment,Text=text.ToString().TrimEnd('\n'),HasBlocks=blocks};
            }
            catch(Refused error){return Reject(error.Message);}
            catch(RegexMatchTimeoutException){return Reject("parse-time-limit");}
            catch(Exception error){return Reject("conversion-failed-"+error.GetType().Name);}
        }
        private static ConversionResult Reject(string status){return new ConversionResult{Status=status};}
        private void Count(int depth){if(depth>MaxDepth)throw new Refused("nesting-limit");if(++nodes>MaxNodes)throw new Refused("node-limit");}
        private void WriteText(string value){html.Append(WebUtility.HtmlEncode(value));text.Append(value);}
        private void Block(Block block,int depth)
        {
            Count(depth);
            var paragraph=block as ParagraphBlock;
            if(paragraph!=null){html.Append("<p>");Inline(paragraph.Inline,depth+1);html.Append("</p>");text.Append('\n');return;}
            var heading=block as HeadingBlock;
            if(heading!=null){if(heading.Level<1||heading.Level>3||heading.IsSetext)throw new Refused("unsupported-heading");formats++;html.Append("<h"+heading.Level+">");Inline(heading.Inline,depth+1);html.Append("</h"+heading.Level+">");text.Append('\n');return;}
            var quote=block as QuoteBlock;
            if(quote!=null){formats++;html.Append("<blockquote>");foreach(var child in quote)Block(child,depth+1);html.Append("</blockquote>");return;}
            var list=block as ListBlock;
            if(list!=null)
            {
                if(depth>0)throw new Refused("nested-list-unverified");
                if(!list.IsOrdered&&list.BulletType!='-')throw new Refused("unsupported-list-marker");
                if(list.IsOrdered&&(list.OrderedDelimiter!='.'||list.OrderedStart!="1"))throw new Refused("ordered-start-unverified");
                formats++;string tag=list.IsOrdered?"ol":"ul";html.Append("<"+tag+">");
                foreach(var child in list)
                {var item=child as ListItemBlock;if(item==null)throw new Refused("unsupported-list-item");Count(depth+1);html.Append("<li>");foreach(var content in item)Block(content,depth+2);html.Append("</li>");}
                html.Append("</"+tag+">");return;
            }
            throw new Refused("unsupported-block-"+block.GetType().Name);
        }
        private void Inline(ContainerInline container,int depth)
        {
            if(container==null)return;
            foreach(var item in container)
            {
                Count(depth);
                var literal=item as LiteralInline;
                if(literal!=null)
                {
                    string value=literal.Content.ToString();
                    // Unescaped dangling subset markers are rejected rather than partially converted.
                    if(!literal.IsFirstCharacterEscaped&&(value.Contains("*")||value.Contains("~~")||value.Contains("`")))throw new Refused("incomplete-inline-markup");
                    WriteText(value);continue;
                }
                var emphasis=item as EmphasisInline;
                if(emphasis!=null)
                {
                    string tag;
                    if(emphasis.DelimiterChar=='*'&&(emphasis.DelimiterCount==1||emphasis.DelimiterCount==2))tag=emphasis.DelimiterCount==1?"i":"b";
                    else if(emphasis.DelimiterChar=='~'&&emphasis.DelimiterCount==2)tag="s";
                    else throw new Refused("unsupported-emphasis");
                    formats++;html.Append("<"+tag+">");Inline(emphasis,depth+1);html.Append("</"+tag+">");continue;
                }
                var code=item as CodeInline;
                if(code!=null){formats++;html.Append("<span style=\"font-family:Consolas;background-color:#eeeeee\">");WriteText(code.Content);html.Append("</span>");continue;}
                var line=item as LineBreakInline;
                if(line!=null){html.Append("<br />");text.Append('\n');continue;}
                if(item.GetType().FullName=="Markdig.Extensions.TaskLists.TaskList")
                {
                    object flag=item.GetType().GetProperty("Checked").GetValue(item,null);
                    formats++;WriteText((bool)flag?"☑ ":"☐ ");continue;
                }
                var entity=item as HtmlEntityInline;
                if(entity!=null)throw new Refused("html-entity-not-in-subset");
                throw new Refused("unsupported-inline-"+item.GetType().Name);
            }
        }
        public static void ValidateHtml(string fragment)
        {
            var settings=new XmlReaderSettings{DtdProcessing=DtdProcessing.Prohibit,XmlResolver=null,MaxCharactersInDocument=MaxOutputBytes};
            var document=new XmlDocument{XmlResolver=null};
            using(var input=new System.IO.StringReader("<root>"+fragment+"</root>"))using(var reader=XmlReader.Create(input,settings))document.Load(reader);
            var allowed=new HashSet<string>{"root","div","p","h1","h2","h3","blockquote","ul","ol","li","i","b","s","span","br"};
            foreach(XmlElement element in document.SelectNodes("//*"))
            {
                if(element.NamespaceURI!=""||!allowed.Contains(element.Name))throw new InvalidOperationException("HTML tag rejected");
                foreach(XmlAttribute attribute in element.Attributes)
                    if(element.Name!="span"||attribute.Name!="style"||attribute.Value!="font-family:Consolas;background-color:#eeeeee")throw new InvalidOperationException("HTML attribute rejected");
            }
        }
    }
}
