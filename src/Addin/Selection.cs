using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Xml;
using OneNote=Microsoft.Office.Interop.OneNote;
namespace OneNoteMarkdown.Addin
{
    internal sealed class Selection
    {
        internal string PageId,Text;
        internal bool WholeParagraphs;
        private byte[] fingerprint;
        private IntPtr root,focus;
        internal bool Same(Selection other)
        {if(other==null||PageId!=other.PageId||root!=other.root||focus!=other.focus||fingerprint.Length!=other.fingerprint.Length)return false;for(int i=0;i<fingerprint.Length;i++)if(fingerprint[i]!=other.fingerprint[i])return false;return true;}
        internal static Selection Read(OneNote.Application app)
        {
            IntPtr root,focus;if(!Native.BodyFocus(out root,out focus))throw new InvalidOperationException("正文未聚焦，请在正文选中后使用快捷键");
            OneNote.Windows windows=null;OneNote.Window window=null;
            try
            {
                windows=app.Windows;if(windows.Count!=1)throw new InvalidOperationException("仅支持单窗口");
                window=windows.CurrentWindow;
                if(window.SideNote||window.DockedLocation!=OneNote.DockLocation.dlNone)throw new InvalidOperationException("此视图尚未支持");
                if(Native.GetAncestor(new IntPtr((long)window.WindowHandle),2)!=root)throw new InvalidOperationException("窗口已变化");
                string xml;app.GetPageContent(window.CurrentPageId,out xml,OneNote.PageInfo.piSelection,OneNote.XMLSchema.xs2013);
                if(xml.Length>1000000)throw new InvalidOperationException("页面过大，已跳过");
                var settings=new XmlReaderSettings{DtdProcessing=DtdProcessing.Prohibit,XmlResolver=null,MaxCharactersInDocument=1000000};
                var document=new XmlDocument{XmlResolver=null};using(var input=new StringReader(xml))using(var reader=XmlReader.Create(input,settings))document.Load(reader);
                const string uri="http://schemas.microsoft.com/office/onenote/2013/onenote";
                if(document.DocumentElement.NamespaceURI!=uri)throw new InvalidOperationException("页面结构未支持");
                var ns=new XmlNamespaceManager(document.NameTable);ns.AddNamespace("o",uri);
                const string selected="[@selected='all' or @selected='partial']";
                if(document.SelectNodes("//o:Title"+selected,ns).Count!=0)throw new InvalidOperationException("不处理标题选区");
                var containers=document.SelectNodes("//o:Outline"+selected,ns);if(containers.Count!=1)throw new InvalidOperationException("选区跨容器或不存在");
                var allowed=new HashSet<string>{"Page","Outline","OEChildren","OE","T"};
                foreach(XmlElement node in document.SelectNodes("//*[@selected]"))
                {string state=node.GetAttribute("selected");if(state=="none")continue;if((state!="all"&&state!="partial")||node.NamespaceURI!=uri||!allowed.Contains(node.LocalName))throw new InvalidOperationException("选区包含表格、对象或未知结构");}
                var paragraphs=document.SelectNodes("//o:OE[o:T[@selected='all']]",ns);
                var text=new StringBuilder();bool whole=true;bool firstLine=true;XmlNode previous=null,parent=null;
                foreach(XmlElement paragraph in paragraphs)
                {
                    if(!paragraph.HasAttribute("objectID"))throw new InvalidOperationException("缺少选区位置标识");
                    if(parent==null)parent=paragraph.ParentNode;
                    if(parent!=paragraph.ParentNode||(previous!=null&&previous.NextSibling!=paragraph))throw new InvalidOperationException("选区不连续");
                    previous=paragraph;
                    if(!ParagraphPolicy.IsPlain(paragraph,document))throw new InvalidOperationException("原生样式或嵌套段落尚未支持");
                    var line=new StringBuilder();bool any=false;
                    foreach(XmlNode child in paragraph.ChildNodes)
                    {
                        if(child.NodeType==XmlNodeType.Whitespace)continue;
                        var run=child as XmlElement;if(run==null||run.LocalName!="T"||run.NamespaceURI!=uri)throw new InvalidOperationException("段落包含非文本结构");
                        // Only the selected run is parsed strictly. Sibling runs (including
                        // styling left by an earlier conversion) survive the paste untouched,
                        // so they must not be parsed or they would refuse a valid second
                        // conversion in the same paragraph.
                        if(run.GetAttribute("selected")=="all")
                        {
                            string value;try{value=TextFragment.Read(run.InnerText);}catch(InvalidOperationException){throw new InvalidOperationException("选中文字已含格式，需为纯文本");}
                            line.Append(value);any=true;
                        }
                        else if(TextFragment.HasContent(run.InnerText))whole=false;
                    }
                    if(any){if(!firstLine)text.Append('\n');text.Append(line);firstLine=false;}
                }
                if(text.Length==0)throw new InvalidOperationException("没有选中文字");
                if(text.Length>MarkdownConverter.MaxCharacters)throw new InvalidOperationException("选区过大");
                IntPtr finalRoot,finalFocus;
                if(!Native.BodyFocus(out finalRoot,out finalFocus)||root!=finalRoot||focus!=finalFocus)throw new InvalidOperationException("焦点已变化");
                byte[] hash;using(var sha=SHA256.Create())hash=sha.ComputeHash(Encoding.UTF8.GetBytes(xml));
                return new Selection{PageId=window.CurrentPageId,Text=text.ToString(),WholeParagraphs=whole,fingerprint=hash,root=root,focus=focus};
            }
            finally{Release(window);Release(windows);}
        }
        internal static void Release(object value){if(value!=null&&Marshal.IsComObject(value))Marshal.ReleaseComObject(value);}
    }
}
