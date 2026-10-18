using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using OneNote = Microsoft.Office.Interop.OneNote;

// Read-only same-page comparison. No UI/input or clipboard access.
internal static class ExportComparison
{
    [STAThread]
    private static int Main()
    {
        bool visible=false;
        foreach(var process in Process.GetProcessesByName("ONENOTE")) using(process) {if(process.MainWindowHandle!=IntPtr.Zero)visible=true;}
        if(!visible){Console.WriteLine("SKIP no visible OneNote; not activating");return 2;}
        OneNote.Application app=null;OneNote.Windows windows=null;OneNote.Window window=null;
        try
        {
            app=new OneNote.Application();windows=app.Windows;
            if(windows.Count!=1){Console.WriteLine("SKIP multiple windows");return 2;}
            window=windows.CurrentWindow;
            string page=window.CurrentPageId,hierarchy;
            app.GetHierarchy(page,OneNote.HierarchyScope.hsSelf,out hierarchy,OneNote.XMLSchema.xs2013);
            if(SelectionScreen.Parse(hierarchy).DocumentElement.GetAttribute("name")!="Markdown插件验证") {Console.WriteLine("SKIP fixture mismatch");return 2;}
            string basic,selected,repeat;
            app.GetPageContent(page,out basic,OneNote.PageInfo.piBasic,OneNote.XMLSchema.xs2013);
            app.GetPageContent(page,out selected,OneNote.PageInfo.piSelection,OneNote.XMLSchema.xs2013);
            app.GetPageContent(page,out repeat,OneNote.PageInfo.piSelection,OneNote.XMLSchema.xs2013);
            if(windows.Count!=1||window.CurrentPageId!=page){Console.WriteLine("SKIP context changed");return 2;}
            Console.WriteLine("Same-page basic versus selection export categories: "+OutsideDiff.Compare(SelectionScreen.Parse(basic),SelectionScreen.Parse(selected)));
            Console.WriteLine("Repeated selection export categories: "+OutsideDiff.Compare(SelectionScreen.Parse(selected),SelectionScreen.Parse(repeat)));
            var document=SelectionScreen.Parse(basic);
            var ns=new System.Xml.XmlNamespaceManager(document.NameTable);ns.AddNamespace("o","http://schemas.microsoft.com/office/onenote/2013/onenote");
            var selectionDocument=SelectionScreen.Parse(selected);
            var selectionNs=new System.Xml.XmlNamespaceManager(selectionDocument.NameTable);selectionNs.AddNamespace("o","http://schemas.microsoft.com/office/onenote/2013/onenote");
            foreach(System.Xml.XmlElement paragraph in selectionDocument.SelectNodes("//o:OE[o:T[@selected='all']]",selectionNs))
            {
                Console.WriteLine("Production paragraph policy accepted="+OneNoteMarkdown.ParagraphPolicy.IsPlain(paragraph,selectionDocument));
                string index=paragraph.GetAttribute("quickStyleIndex");
                System.Xml.XmlElement definition=null;
                foreach(System.Xml.XmlElement candidate in selectionDocument.SelectNodes("//o:QuickStyleDef",selectionNs))if(candidate.GetAttribute("index")==index)definition=candidate;
                Console.WriteLine("Selected paragraph diagnostic: quickStyle="+paragraph.HasAttribute("quickStyleIndex")+" definitionFound="+(definition!=null)+" list="+(paragraph.SelectSingleNode("o:List",selectionNs)!=null)+" tag="+(paragraph.SelectSingleNode("o:Tag",selectionNs)!=null)+" nested="+(paragraph.SelectSingleNode("o:OEChildren",selectionNs)!=null));
                if(definition!=null)
                {
                    string standardName=definition.GetAttribute("name");
                    Console.WriteLine("QuickStyle known normal name="+(standardName=="Normal"||standardName=="正文"||standardName=="普通"||standardName=="普通文本")+" paragraph-name="+(standardName=="p")+" attributes="+String.Join(",",System.Linq.Enumerable.Select(System.Linq.Enumerable.Cast<System.Xml.XmlAttribute>(definition.Attributes),a=>a.LocalName)));
                }
            }
            int matches=0;
            foreach(System.Xml.XmlElement text in document.SelectNodes("//o:T",ns))
            {
                System.Xml.XmlDocument fragment;
                try{fragment=FixtureText.Parse(text.InnerText);}catch(System.Xml.XmlException){continue;}
                if(fragment.DocumentElement.InnerText!="code")continue;
                matches++;
                bool mono=false;
                foreach(System.Xml.XmlElement span in fragment.SelectNodes("//*"))
                {
                    string style=span.GetAttribute("style");
                    if(style.IndexOf("Consolas",StringComparison.OrdinalIgnoreCase)>=0||style.IndexOf("Courier",StringComparison.OrdinalIgnoreCase)>=0||style.IndexOf("monospace",StringComparison.OrdinalIgnoreCase)>=0)mono=true;
                }
                bool ancestorMono=false;
                for(System.Xml.XmlNode ancestor=text.ParentNode;ancestor is System.Xml.XmlElement;ancestor=ancestor.ParentNode)
                {
                    string style=((System.Xml.XmlElement)ancestor).GetAttribute("style");
                    if(style.IndexOf("Consolas",StringComparison.OrdinalIgnoreCase)>=0||style.IndexOf("Courier",StringComparison.OrdinalIgnoreCase)>=0||style.IndexOf("monospace",StringComparison.OrdinalIgnoreCase)>=0)
                    {
                        ancestorMono=true;
                        Console.WriteLine("Code monospace ancestor kind="+ancestor.LocalName+" explicit Consolas="+(style.IndexOf("Consolas",StringComparison.OrdinalIgnoreCase)>=0));
                    }
                }
                Console.WriteLine("Exact code fixture text: monospace declared in fragment="+mono+" ancestor="+ancestorMono);
            }
            Console.WriteLine("Exact code fixture matches="+matches+"; ancestor/default styles not resolved");
            int alphaParagraphs=0,unicodeParagraphs=0,exactBlocks=0;
            foreach(System.Xml.XmlElement outline in document.SelectNodes("//o:Outline",ns))
            {
                var paragraphs=new System.Collections.Generic.List<string>();
                foreach(System.Xml.XmlElement oe in outline.SelectNodes(".//o:OE[o:T]",ns))
                {
                    var text=new System.Text.StringBuilder();bool decodable=true;
                    foreach(System.Xml.XmlElement run in oe.SelectNodes("o:T",ns))
                    {
                        try
                        {
                            var fragment=FixtureText.Parse(run.InnerText);
                            foreach(System.Xml.XmlElement br in fragment.SelectNodes("//br"))br.ParentNode.ReplaceChild(fragment.CreateTextNode("\n"),br);
                            text.Append(fragment.DocumentElement.InnerText);
                        }
                        catch(System.Xml.XmlException){decodable=false;break;}
                    }
                    if(!decodable)continue;
                    string value=text.ToString().Replace('\u00a0',' ').Replace("\r\n","\n");
                    paragraphs.Add(value);
                    if(value.Contains("alpha")) {alphaParagraphs++;Console.WriteLine("Code alpha fixture has four-space prefix="+(value.StartsWith("    alpha",StringComparison.Ordinal))+" decoded chars="+value.Length+" newline count="+(value.Length-value.Replace("\n","").Length));}
                    if(value.Contains("中文😀")){unicodeParagraphs++;int at=value.IndexOf("中文😀",StringComparison.Ordinal),spaces=0;for(int k=at-1;k>=0&&value[k]==' ';k--)spaces++;Console.WriteLine("Code Unicode fixture has four-space line="+(value.StartsWith("    中文😀",StringComparison.Ordinal)||value.Contains("\n    中文😀"))+" immediately preceding spaces="+spaces+" decoded chars="+value.Length);}
                }
                string combined=String.Join("\n",paragraphs);
                if(combined.Contains("    alpha\n\n    中文😀"))exactBlocks++;
            }
            Console.WriteLine("Fenced fixture matching outlines="+exactBlocks+" alpha paragraph candidates="+alphaParagraphs+" Unicode paragraph candidates="+unicodeParagraphs+"; NBSP treated as visual space");
            Console.WriteLine("No document edits or clipboard access; not a before/after paste comparison");
            return 0;
        }
        catch(Exception ex){Console.WriteLine("FAIL "+ex.GetType().Name+" HRESULT=0x"+ex.HResult.ToString("X8"));return 1;}
        finally{foreach(object obj in new object[]{window,windows,app})if(obj!=null&&Marshal.IsComObject(obj))Marshal.ReleaseComObject(obj);}
    }
}
