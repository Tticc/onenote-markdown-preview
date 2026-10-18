using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Xml;
using OneNote = Microsoft.Office.Interop.OneNote;

// One authorized inline fixture only. Human acknowledgement is required before restore.
internal static class PasteFixture
{
    [StructLayout(LayoutKind.Sequential)] private struct Rect { public int A,B,C,D; }
    [StructLayout(LayoutKind.Sequential)] private struct Gui { public int Size,Flags; public IntPtr Active,Focus,Capture,Menu,Move,Caret; public Rect R; }
    [StructLayout(LayoutKind.Sequential)] private struct Input { public uint Type; public Union U; }
    [StructLayout(LayoutKind.Explicit)] private struct Union { [FieldOffset(0)] public Keyboard K; [FieldOffset(0)] public Mouse M; }
    [StructLayout(LayoutKind.Sequential)] private struct Keyboard { public ushort Key,Scan; public uint Flags,Time; public UIntPtr Extra; }
    [StructLayout(LayoutKind.Sequential)] private struct Mouse { public int X,Y; public uint Data,Flags,Time; public UIntPtr Extra; }
    [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] private static extern IntPtr GetAncestor(IntPtr w,uint f);
    [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr w,out uint p);
    [DllImport("user32.dll")] private static extern bool GetGUIThreadInfo(uint t,ref Gui g);
    [DllImport("user32.dll",CharSet=CharSet.Unicode)] private static extern int GetClassName(IntPtr w,StringBuilder n,int c);
    [DllImport("user32.dll")] private static extern bool OpenClipboard(IntPtr w);
    [DllImport("user32.dll")] private static extern bool CloseClipboard();
    [DllImport("user32.dll")] private static extern uint GetClipboardSequenceNumber();
    [DllImport("user32.dll")] private static extern IntPtr GetClipboardOwner();
    [DllImport("user32.dll",CharSet=CharSet.Unicode)] private static extern uint RegisterClipboardFormat(string n);
    [DllImport("user32.dll")] private static extern uint SendInput(uint n,Input[] i,int size);
    [DllImport("user32.dll")] private static extern short GetAsyncKeyState(int key);
    private static string log;
    private static bool blockFixture;
    private static bool codeFixture;
    private static bool fencedFixture;
    private static bool explicitWhitespace;
    private static bool rtfFixture;
    private static bool remainingFixture;
    internal const string CodeHtml="<span style=\"font-family:Consolas;background-color:#eeeeee\">code</span>";
    private static Input[] PasteInputs()
    {
        var input=new Input[4];
        for(int i=0;i<input.Length;i++) input[i].Type=1; // INPUT_KEYBOARD, not default INPUT_MOUSE.
        input[0].U.K.Key=0x11;input[1].U.K.Key=0x56;
        input[2].U.K.Key=0x56;input[2].U.K.Flags=2;
        input[3].U.K.Key=0x11;input[3].U.K.Flags=2;
        return input;
    }
    private static void Log(string message) { File.AppendAllText(log,DateTime.UtcNow.ToString("o")+" "+message+Environment.NewLine); }
    private static void Observe(OneNote.Application app,string before,string basicBefore)
    {
        OneNote.Windows windows=null;OneNote.Window window=null;
        try
        {
            var original=SelectionScreen.Parse(before);
            string pageId=original.DocumentElement.GetAttribute("ID");
            if(String.IsNullOrEmpty(pageId)) {Log("OBSERVATION unavailable page identity");return;}
            windows=app.Windows;
            if(windows.Count!=1) {Log("OBSERVATION canceled window count");return;}
            window=windows.CurrentWindow;
            var watch=Stopwatch.StartNew();
            string previousReason=null;
            while(watch.ElapsedMilliseconds<2500)
            {
                if(windows.Count!=1 || window.CurrentPageId!=pageId) {Log("OBSERVATION canceled page/window change");return;}
                string after;
                app.GetPageContent(pageId,out after,OneNote.PageInfo.piBasic,OneNote.XMLSchema.xs2013);
                string reason=InlinePostcondition.Diagnose(before,after,basicBefore);
                if(reason!=previousReason) {Log("OBSERVATION category="+reason);previousReason=reason;}
                if(reason=="expected-text-and-outside-serialization-match") {Log("OBSERVATION expected inline text and outside content match; not style/transaction proof");return;}
                Thread.Sleep(50);
            }
            Log("OBSERVATION expected postcondition not confirmed; no retry/undo");
        }
        catch(Exception ex) {Log("OBSERVATION failed "+ex.GetType().Name+"; no retry/undo");}
        finally {foreach(object obj in new object[]{window,windows})if(obj!=null&&Marshal.IsComObject(obj))Marshal.ReleaseComObject(obj);}
    }
    private static SelectionEvidence Capture(OneNote.Application app, out string xml)
    {
        xml=null;
        OneNote.Windows windows=null; OneNote.Window window=null;
        try
        {
            windows=app.Windows;
            if(windows.Count!=1) return null;
            window=windows.CurrentWindow;
            if(window.SideNote) return null;
            string hierarchy;
            app.GetHierarchy(window.CurrentPageId,OneNote.HierarchyScope.hsSelf,out hierarchy,OneNote.XMLSchema.xs2013);
            Log("CHECK hierarchy metadata");
            if(SelectionScreen.Parse(hierarchy).DocumentElement.GetAttribute("name")!="Markdown插件验证") return null;
            app.GetPageContent(window.CurrentPageId,out xml,OneNote.PageInfo.piSelection,OneNote.XMLSchema.xs2013);
            Log("CHECK selection XML");
            var doc=SelectionScreen.Parse(xml);
            var ns=new XmlNamespaceManager(doc.NameTable); ns.AddNamespace("o","http://schemas.microsoft.com/office/onenote/2013/onenote");
            var nodes=doc.SelectNodes("//o:T[@selected='all']",ns);
            if(blockFixture||fencedFixture||remainingFixture)
            {
                if(!BlockFixture.Matches(xml,remainingFixture?BlockFixture.RemainingSource:(fencedFixture?BlockFixture.FencedSource:BlockFixture.Source)))return null;
                Log("CHECK exact block fixture and whole paragraphs");
            }
            else
            {
                if(nodes.Count!=1) return null;
                Log("CHECK selected text fragment");
                var fragment=FixtureText.Parse(nodes[0].InnerText);
                if(fragment.DocumentElement.InnerText!=(codeFixture?"`code`":"**测试文字**")) return null;
            }
            uint pid; uint thread=GetWindowThreadProcessId(new IntPtr((long)window.WindowHandle),out pid);
            var gui=new Gui{Size=Marshal.SizeOf(typeof(Gui))};
            if(!GetGUIThreadInfo(thread,ref gui)) return null;
            var name=new StringBuilder(256); GetClassName(gui.Focus,name,256);
            return SelectionEvidence.Capture(xml,window.CurrentPageId,GetAncestor(new IntPtr((long)window.WindowHandle),2).ToInt64(),gui.Focus.ToInt64(),GetAncestor(GetForegroundWindow(),2).ToInt64(),name.ToString(),(int)windows.Count);
        }
        finally { foreach(object obj in new object[]{window,windows}) if(obj!=null&&Marshal.IsComObject(obj)) Marshal.ReleaseComObject(obj); }
    }
    [STAThread] private static int Main(string[] args)
    {
        if(args.Length==1 && args[0]=="--check-input-layout")
        {
            var inputs=PasteInputs();
            if(Marshal.SizeOf(typeof(Input))!=40 || Marshal.OffsetOf(typeof(Input),"U").ToInt32()!=8) return 1;
            for(int i=0;i<4;i++)
            {
                if(inputs[i].Type!=1 || inputs[i].U.K.Key!=(i==0||i==3?0x11:0x56) || inputs[i].U.K.Flags!=(i>=2?2u:0u)) return 1;
            }
            Console.WriteLine("PASS x64 INPUT layout and four keyboard Ctrl+V events; no input sent");
            return 0;
        }
        bool reset=false;int delay=-1;
        if(args.Length<2)return 2;
        for(int i=2;i<args.Length;i++)
        {
            if(args[i]=="--reset-disposable")reset=true;
            else if(args[i]=="--delay=200")delay=200;
            else if(args[i]=="--delay=500")delay=500;
            else if(args[i]=="--block-fixture")blockFixture=true;
            else if(args[i]=="--code-fixture")codeFixture=true;
            else if(args[i]=="--fenced-fixture")fencedFixture=true;
            else if(args[i]=="--explicit-whitespace")explicitWhitespace=true;
            else if(args[i]=="--rtf-code")rtfFixture=true;
            else if(args[i]=="--remaining-fixture")remainingFixture=true;
            else return 2;
        }
        if((codeFixture?1:0)+(blockFixture?1:0)+(fencedFixture?1:0)+(remainingFixture?1:0)>1)return 2;
        if(explicitWhitespace){return 2;} // NBSP candidate rejected by user's strict whitespace requirement.
        if(rtfFixture&&!fencedFixture)return 2;
        log=args[0];
        OneNote.Application app=null;
        Dictionary<uint,IntPtr> backup=null,payload=null;
        using(var done=new EventWaitHandle(false,EventResetMode.ManualReset,args[1]))
        using(var owner=new Owner())
        try
        {
            Log("WAIT 12 seconds; keep fixture selected in OneNote"); Thread.Sleep(12000);
            bool visible=false; foreach(var p in Process.GetProcessesByName("ONENOTE")) using(p) { if(p.MainWindowHandle!=IntPtr.Zero) visible=true; }
            if(!visible) { Log("SKIP no visible host"); return 2; }
            app=new OneNote.Application();
            string xml; var initial=Capture(app,out xml);
            if(initial==null) { Log("SKIP initial selection/focus/fixture guard"); return 2; }
            uint original;
            Dictionary<uint,byte[]> savedSnapshot=null;
            if(reset)
            {
                var fixture=ClipboardRoundtrip.Allocate(new Dictionary<uint,byte[]>{{13,Encoding.Unicode.GetBytes("clipboard-probe-123\0")}});
                try
                {
                    if(!OpenClipboard(owner.Handle)) {Log("SKIP fixture reset busy");return 2;}
                    try {ClipboardRoundtrip.Write(fixture);Log("Reset current clipboard to explicitly authorized disposable fixture");}
                    finally {CloseClipboard();}
                }
                finally {ClipboardRoundtrip.Free(fixture);}
            }
            if(!OpenClipboard(owner.Handle)) { Log("SKIP clipboard busy"); return 2; }
            try
            {
                var snapshot=ClipboardRoundtrip.Read();
                savedSnapshot=snapshot;
                byte[] text;
                if(!snapshot.TryGetValue(13,out text)||Encoding.Unicode.GetString(text).TrimEnd('\0')!="clipboard-probe-123") { Log("SKIP disposable clipboard mismatch"); return 2; }
                backup=ClipboardRoundtrip.Allocate(snapshot); original=GetClipboardSequenceNumber();
            }
            finally { CloseClipboard(); }
            original=GetClipboardSequenceNumber();
            var contentPayload=new Dictionary<uint,byte[]>();
            if(rtfFixture)contentPayload.Add(RegisterClipboardFormat("Rich Text Format"),Encoding.ASCII.GetBytes(RtfCodeFixture.Render("    alpha\n\n    中文😀")+"\0"));
            else contentPayload.Add(RegisterClipboardFormat("HTML Format"),HtmlFixturePayload.Create(remainingFixture?BlockFixture.RemainingHtml:(fencedFixture?BlockFixture.FencedHtml:(blockFixture?BlockFixture.Html:(codeFixture?CodeHtml:"<span style=\"font-weight:bold\">测试文字</span>")))));
            contentPayload.Add(13,Encoding.Unicode.GetBytes((remainingFixture?"斜体 删除线\r\n1. 有序项\r\n☐ 未完成\r\n☑ 已完成":(fencedFixture?"    alpha\r\n\r\n    中文😀":(blockFixture?"测试标题\r\n第一项\r\n引用文字\r\ncode":(codeFixture?"code":"测试文字"))))+"\0"));
            payload=ClipboardRoundtrip.Allocate(contentPayload);
            string second; if(!initial.Matches(Capture(app,out second))) { Log("SKIP evidence changed"); return 2; }
            var beforeDocument=SelectionScreen.Parse(xml);
            string basicBefore;
            app.GetPageContent(beforeDocument.DocumentElement.GetAttribute("ID"),out basicBefore,OneNote.PageInfo.piBasic,OneNote.XMLSchema.xs2013);
            if(delay<0 && !InlinePostcondition.BaselineMatches(xml,basicBefore)) {Log("SKIP baseline target differs across export modes before paste");return 2;}
            if(!initial.Matches(Capture(app,out second))) {Log("SKIP evidence changed during baseline export");return 2;}
            foreach(int key in new[]{0x10,0x11,0x12,0x5B,0x5C}) if(GetAsyncKeyState(key)<0) { Log("SKIP modifier held"); return 2; }
            if(!OpenClipboard(owner.Handle)) { Log("SKIP clipboard publish busy"); return 2; }
            try
            {
                if(GetClipboardSequenceNumber()!=original) {Log("SKIP clipboard changed");return 2;}
                var current=ClipboardRoundtrip.Read();
                bool equivalent=current.Count==savedSnapshot.Count;
                foreach(var pair in savedSnapshot)
                {
                    byte[] value;
                    if(!current.TryGetValue(pair.Key,out value)||!Convert.ToBase64String(value).Equals(Convert.ToBase64String(pair.Value))) equivalent=false;
                }
                if(!equivalent) {Log("SKIP clipboard snapshot differs");return 2;}
                ClipboardRoundtrip.Write(payload);
            }
            finally {CloseClipboard();}
            uint published=GetClipboardSequenceNumber();
            bool same=initial.Matches(Capture(app,out second));
            if(same)
            {
                var input=PasteInputs();
                uint sent=SendInput(4,input,Marshal.SizeOf(typeof(Input)));
                Log("PASTE_SUBMITTED inputCount="+sent+"; no automatic retry/undo; mode="+(delay>=0?"empirical-delay":"diagnostic-acknowledgement"));
                if(delay>=0)
                {
                    Log("Configured empirical restoration delay="+RestorePolicy.Delay(delay)+"ms; not completion proof");
                    Thread.Sleep(RestorePolicy.Delay(delay));
                }
                else
                {
                    if(sent==4) Observe(app,xml,basicBefore);
                    if(!done.WaitOne(300000)) { Log("UNCONFIRMED timeout; no automatic restore while paste outcome unknown"); return 3; }
                    Log("Diagnostic acknowledgement received; attempting clipboard restore");
                }
            }
            else Log("CANCEL before input; restore payload if still owned");
            if(!OpenClipboard(owner.Handle)) {Log("RESTORE_FAILED busy; disposable clipboard may retain payload");return 3;}
            try
            {
                if(!RestorePolicy.MayRestore(owner.Handle.ToInt64(),GetClipboardOwner().ToInt64(),published,GetClipboardSequenceNumber())) {Log("CHANGED_CLIPBOARD preserved; snapshot not restored");return 0;}
                ClipboardRoundtrip.Write(backup);Log("RESTORED original disposable clipboard formats");
            }
            finally {CloseClipboard();}
            return 0;
        }
        catch(Exception ex) {Log("FAIL "+ex.GetType().Name+" HRESULT=0x"+ex.HResult.ToString("X8")+"; no retry/undo");return 1;}
        finally {ClipboardRoundtrip.Free(payload);ClipboardRoundtrip.Free(backup);if(app!=null&&Marshal.IsComObject(app))Marshal.ReleaseComObject(app);}
    }
    private sealed class Owner:NativeWindow,IDisposable
    { internal Owner(){CreateHandle(new CreateParams{Parent=new IntPtr(-3)});}public void Dispose(){DestroyHandle();} }
}
