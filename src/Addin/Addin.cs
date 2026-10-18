using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Reflection;
using System.Windows.Forms;
using Extensibility;
using Microsoft.Office.Core;
using OneNote=Microsoft.Office.Interop.OneNote;
namespace OneNoteMarkdown.Addin
{
    [ComVisible(true),Guid("E036DE73-418B-45F7-973F-A4B09D58E515"),ProgId("OneNoteMarkdown.Preview"),ClassInterface(ClassInterfaceType.AutoDispatch)]
    public sealed class Addin:IDTExtensibility2,IRibbonExtensibility
    {
        private Thread thread;private Control dispatcher;private ApplicationContext context;
        private readonly OperationState state=new OperationState();
        private Native.Hook hookCallback;private IntPtr hook;
        private bool keyDown,suppressed;
        private string status="未启用（预览版）";
        private ClipboardSession retained;
        private string folder;
        private volatile bool shuttingDown;
        private readonly bool offlineLifecycle;
        private readonly string testFolder;
        private bool lifecycleOnly;
        private Toast toast;
        private int workerFinalized;private int exitTraced;
        public Addin(){}
        internal Addin(string folder){testFolder=folder;offlineLifecycle=true;}
        private static readonly ResolveEventHandler resolver=ResolveDependency;
        private static readonly System.DateTime SurrogateStart=ReadProcessStart();
        private static System.DateTime ReadProcessStart()
        {
            try{return System.Diagnostics.Process.GetCurrentProcess().StartTime.ToUniversalTime();}
            catch(System.Exception){return System.DateTime.UtcNow;}
        }
        private void TraceLifecycle(string stage)
        {
            try
            {
                if(folder==null)return;
                Directory.CreateDirectory(folder);
                // Elapsed since the surrogate process started isolates this add-in's own load
                // cost from the host, which the startup samples cannot separate.
                string elapsed="+"+(long)(System.DateTime.UtcNow-SurrogateStart).TotalMilliseconds+"ms-since-surrogate-start";
                File.AppendAllText(Path.Combine(folder,"lifecycle.log"),DateTime.UtcNow.ToString("o")+" pid="+System.Diagnostics.Process.GetCurrentProcess().Id+" "+elapsed+" "+stage+Environment.NewLine);
            }
            catch(Exception){}
        }
        static Addin(){AppDomain.CurrentDomain.AssemblyResolve+=resolver;}
        public void OnConnection(object application,ext_ConnectMode mode,object addin,ref Array custom)
        {
            // Never retain incoming application/AddInInst COM proxies.
            if(thread!=null&&thread.IsAlive)return;
            if(shuttingDown)return;
            folder=testFolder??Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),"OneNoteMarkdownPreview");
            // Lifecycle logging on the worker avoids filesystem work in host callbacks.
            thread=new Thread(Run){IsBackground=true,Name="OneNoteMarkdown STA"};thread.SetApartmentState(ApartmentState.STA);thread.Start();
        }
        private void Run()
        {
            TraceLifecycle("addin-connected");
            try
            {
                Application.OleRequired();
                using(var control=new Control())
                using(var stopTimer=new System.Windows.Forms.Timer{Interval=50})
                {
                    control.CreateControl();context=new ApplicationContext();dispatcher=control;
                    // Host shutdown callbacks never wait for this thread. The flag also
                    // covers disconnect arriving before a handle/message loop exists.
                    stopTimer.Tick+=delegate{if(shuttingDown){state.Disable();Unhook();if(!state.Busy)context.ExitThread();}};
                    Directory.CreateDirectory(folder);
                    TraceLifecycle("worker-started");
                    lifecycleOnly=File.Exists(Path.Combine(folder,"lifecycle-only"));
                    bool unfinished=File.Exists(Path.Combine(folder,"operation.pending"));
                    if(lifecycleOnly)SetStatus("生命周期诊断：转换与快捷键锁定");
                    else if(!shuttingDown&&!offlineLifecycle&&!unfinished&&File.Exists(Path.Combine(folder,"enabled")))Enable();
                    else if(unfinished){state.Fault();SetStatus("上次操作未完成，需手动启用");}
                    if(!shuttingDown){stopTimer.Start();Application.Run(context);}
                }
            }
            catch(Exception){state.Fault();SetStatus("入口异常，需重新连接");}
            finally{
                // Idempotent: a CLR shutdown/thread-abort during disconnect can re-enter this
                // finally, which would double-run cleanup and duplicate the trace line.
                if(Interlocked.Exchange(ref workerFinalized,1)==0){
                    Unhook();if(retained!=null){retained.Dispose();retained=null;}if(toast!=null){try{if(!toast.IsDisposed)toast.Close();}catch(Exception){}toast=null;}dispatcher=null;if(context!=null){context.Dispose();context=null;}TraceExited();
                }
            }
        }
        private static Assembly ResolveDependency(object sender,ResolveEventArgs args)
        {
            var name=new AssemblyName(args.Name);
            string[] approved={"OneNoteMarkdown.Core","Markdig","System.Memory","System.Buffers","System.Numerics.Vectors","System.Runtime.CompilerServices.Unsafe"};
            if(Array.IndexOf(approved,name.Name)<0)return null;
            string path=Path.Combine(Path.GetDirectoryName(new Uri(typeof(Addin).Assembly.CodeBase).LocalPath),name.Name+".dll");
            if(!File.Exists(path))return null;
            var actual=AssemblyName.GetAssemblyName(path);
            if(actual.Name!=name.Name||(name.Version!=null&&actual.Version!=name.Version))return null;
            return Assembly.LoadFrom(path);
        }
        private void Post(Action action)
        {try{var target=dispatcher;if(!shuttingDown&&target!=null&&!target.IsDisposed)target.BeginInvoke(action);}catch(InvalidOperationException){}}
        public string GetCustomUI(string ribbonId)
        {return "<customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\"><ribbon><tabs><tab id=\"SafeMdTab\" label=\"Markdown 预览\"><group id=\"SafeMdGroup\" label=\"选区格式化\"><button id=\"SafeMdEnable\" label=\"启用\" onAction=\"EnableClicked\"/><button id=\"SafeMdDisable\" label=\"禁用\" onAction=\"DisableClicked\"/><button id=\"SafeMdFormat\" label=\"格式化选区\" onAction=\"FormatClicked\" screentip=\"在正文选中后使用 Ctrl+Shift+D；不强行恢复焦点\"/><button id=\"SafeMdRecover\" label=\"重试恢复剪贴板\" onAction=\"RecoverClicked\"/><labelControl id=\"SafeMdStatus\" label=\"最近结果见本地 status.txt\"/></group></tab></tabs></ribbon></customUI>";}
        // Compatibility entry only: never retain this host-owned COM interface.
        public void RibbonLoaded(IRibbonUI ui){}
        public string GetStatus(IRibbonControl control){return status;}
        public void EnableClicked(IRibbonControl control){Post(Enable);}
        public void DisableClicked(IRibbonControl control){Post(Disable);}
        public void FormatClicked(IRibbonControl control){Post(()=>Format());}
        public void RecoverClicked(IRibbonControl control){Post(Recover);}
        private void SetStatus(string value)
        {
            status=value;
            // No synchronous callback into Office from our STA, including during
            // disconnect. Local bounded status replaces Ribbon invalidation for now.
            try{if(folder!=null)File.WriteAllText(Path.Combine(folder,"status.txt"),value);}catch(Exception){}
        }
        private void Alert(string value)
        {
            SetStatus(value);
            // Non-activating, auto-dismissing popup for a skip/failure. Never a modal
            // dialog, never steals focus, never retained across disconnect. Offline
            // fixtures and shutdown skip it. Message is metadata only (reason and
            // clipboard format name), never note text, HTML or clipboard bytes.
            if(offlineLifecycle||shuttingDown)return;
            try
            {
                var target=dispatcher;if(target==null||target.IsDisposed)return;
                if(toast!=null){try{if(!toast.IsDisposed)toast.Close();}catch(Exception){}toast=null;}
                var next=new Toast(Notice(value));
                var area=Screen.PrimaryScreen.WorkingArea;
                next.Location=new System.Drawing.Point(area.Right-next.Width-16,area.Bottom-next.Height-16);
                next.Show();
                toast=next;
            }
            catch(Exception){}
        }
        // Popup text only: status.txt keeps the raw reason code for diagnostics, while the
        // popup shows a plain-Chinese explanation of the same reason.
        private static readonly Dictionary<string,string> SkipReasons=new Dictionary<string,string>(StringComparer.Ordinal)
        {
            {"no-markdown","选中内容不含 Markdown"},
            {"fenced-code-fidelity-unverified","代码块不在本版本支持范围（计划后续版本）"},
            {"whole-paragraphs-required","需要选中完整段落"},
            {"unsupported-heading","标题层级不支持"},
            {"unsupported-emphasis","强调符号不支持"},
            {"unsupported-list-marker","列表符号不支持"},
            {"unsupported-list-item","列表项结构不支持"},
            {"input-limit","选中内容过长"},
            {"invalid-input","选中内容无效"},
            {"invalid-character","选中内容含无效字符"},
            {"parse-time-limit","解析超时"},
            {"conversion-failed","转换失败"}
        };
        private static string DescribeSkip(string code)
        {
            string text;
            if(SkipReasons.TryGetValue(code,out text))return text;
            if(code.StartsWith("unsupported-",StringComparison.Ordinal))return "含不支持的语法";
            if(code.StartsWith("conversion-failed-",StringComparison.Ordinal))return "转换失败";
            return code;
        }
        private static string Notice(string value)
        {
            if(value.StartsWith("跳过：",StringComparison.Ordinal))return "跳过："+DescribeSkip(value.Substring(3));
            return value;
        }
        private sealed class Toast:Form
        {
            private readonly System.Windows.Forms.Timer life;
            internal Toast(string value)
            {
                FormBorderStyle=FormBorderStyle.None;ShowInTaskbar=false;TopMost=true;StartPosition=FormStartPosition.Manual;
                BackColor=System.Drawing.Color.FromArgb(250,250,245);
                ClientSize=new System.Drawing.Size(380,88);
                var title=new Label{Text="OneNote Markdown 预览",Font=new System.Drawing.Font("Microsoft YaHei UI",9F,System.Drawing.FontStyle.Bold),AutoSize=false,Location=new System.Drawing.Point(12,8),Size=new System.Drawing.Size(356,20),ForeColor=System.Drawing.Color.FromArgb(60,60,60)};
                var body=new Label{Text=value,Font=new System.Drawing.Font("Microsoft YaHei UI",9F,System.Drawing.FontStyle.Regular),AutoSize=false,Location=new System.Drawing.Point(12,32),Size=new System.Drawing.Size(356,48),ForeColor=System.Drawing.Color.FromArgb(40,40,40)};
                Controls.Add(title);Controls.Add(body);
                life=new System.Windows.Forms.Timer{Interval=6000};
                life.Tick+=delegate{life.Stop();Close();};
            }
            protected override bool ShowWithoutActivation{get{return true;}}
            protected override CreateParams CreateParams{get{var parameters=base.CreateParams;parameters.ExStyle|=0x08000080;return parameters;}}
            protected override void OnShown(EventArgs e){base.OnShown(e);life.Start();}
            protected override void Dispose(bool disposing){if(disposing){life.Dispose();}base.Dispose(disposing);}
        }
        private void Persist(bool enabled)
        {string file=Path.Combine(folder,"enabled");if(enabled)File.WriteAllText(file,"1");else if(File.Exists(file))File.Delete(file);}
        private void Enable()
        {
            if(lifecycleOnly){SetStatus("生命周期诊断：禁止启用转换");return;}
            if(shuttingDown||offlineLifecycle)return;
            if(state.Enabled&&hook!=IntPtr.Zero){SetStatus("已启用：Ctrl+Shift+D");return;}
            if(retained!=null){Alert("先处理待恢复快照");return;}
            if(!state.Enable()){Alert("前次任务尚未结束");return;}
            try
            {
                Persist(true);string pending=Path.Combine(folder,"operation.pending");if(File.Exists(pending))File.Delete(pending);
                hookCallback=Keyboard;hook=Native.SetWindowsHookEx(13,hookCallback,Native.GetModuleHandle(null),0);
                if(hook==IntPtr.Zero)throw new InvalidOperationException();SetStatus("已启用：Ctrl+Shift+D");
            }
            catch(Exception){state.Fault();Unhook();Alert("启用失败");}
        }
        private void Disable(){state.Disable();Unhook();try{Persist(false);}catch(Exception){}SetStatus("已禁用");}
        private void Unhook(){if(hook!=IntPtr.Zero){Native.UnhookWindowsHookEx(hook);hook=IntPtr.Zero;}keyDown=false;suppressed=false;}
        private IntPtr Keyboard(int code,IntPtr message,IntPtr data)
        {
            try
            {
                if(code>=0&&Marshal.ReadInt32(data)==0x44&&(Marshal.ReadInt32(data,8)&0x10)==0)
                {
                    int msg=message.ToInt32();
                    if(msg==0x101||msg==0x105){bool eat=suppressed;keyDown=false;suppressed=false;if(eat)return new IntPtr(1);}
                    else if(msg==0x100||msg==0x104)
                    {
                        if(keyDown)return suppressed?new IntPtr(1):Native.CallNextHookEx(hook,code,message,data);
                        keyDown=true;IntPtr root,focus;
                        if(state.Enabled&&Native.GetAsyncKeyState(0x11)<0&&Native.GetAsyncKeyState(0x10)<0&&Native.GetAsyncKeyState(0x12)>=0&&Native.BodyFocus(out root,out focus))
                        {suppressed=true;Post(()=>Format());return new IntPtr(1);}
                    }
                }
            }
            catch(Exception){}
            return Native.CallNextHookEx(hook,code,message,data);
        }
        private void Recover()
        {
            if(retained==null){Alert("无待恢复快照");return;}
            try{string outcome=retained.Restore();if(outcome=="剪贴板已恢复")SetStatus(outcome);else Alert(outcome);if(!retained.Published){retained.Dispose();retained=null;string pending=Path.Combine(folder,"operation.pending");if(File.Exists(pending))File.Delete(pending);}}
            catch(Exception){Alert("恢复失败，保留内存快照；不重试粘贴");}
        }
        private async void Format()
        {
            if(lifecycleOnly||shuttingDown)return;
            long ticket=state.Begin();if(ticket==0)return;
            OneNote.Application app=null;ClipboardSession clipboard=null;bool parserOutstanding=false;bool submitted=false;
            string pending=Path.Combine(folder,"operation.pending");
            try
            {
                if(retained!=null)throw new InvalidOperationException("先处理待恢复剪贴板");
                // Wait for physical shortcut modifiers to be released, never synthesize key-up.
                var watch=System.Diagnostics.Stopwatch.StartNew();
                while(Native.GetAsyncKeyState(0x10)<0||Native.GetAsyncKeyState(0x11)<0)
                {if(!state.Valid(ticket))return;if(watch.ElapsedMilliseconds>1000)throw new InvalidOperationException("请松开快捷键后重试");await Task.Delay(20);}
                if(!state.Valid(ticket))return;
                app=new OneNote.Application();var selection=Selection.Read(app);
                File.WriteAllText(pending,"pending");
                var clock=System.Diagnostics.Stopwatch.StartNew();
                Task<ConversionResult> conversion=Task.Run(()=>new MarkdownConverter().Convert(selection.Text,selection.WholeParagraphs));
                if(await Task.WhenAny(conversion,Task.Delay(2000))!=conversion)
                {
                    state.Fault();Persist(false);Unhook();parserOutstanding=true;Alert("转换超时，已停用");
                    RegisterLateCompletion(conversion);return;
                }
                var result=await conversion;TraceLifecycle("conversion-ms="+clock.ElapsedMilliseconds+" chars="+selection.Text.Length);
                if(!state.Valid(ticket)){Alert("操作已取消");return;}
                if(!result.CanApply){Alert("跳过："+result.Status);return;}
                clipboard=new ClipboardSession();clipboard.Capture();
                if(!state.Valid(ticket)||!selection.Same(Selection.Read(app))){Alert("选区变化，已取消");return;}
                clipboard.Publish(result);
                if(!state.Valid(ticket)||!selection.Same(Selection.Read(app))){Alert(clipboard.Restore());return;}
                foreach(int key in new[]{0x10,0x11,0x12,0x5B,0x5C})if(Native.GetAsyncKeyState(key)<0)throw new InvalidOperationException("检测到修饰键，已取消");
                submitted=true;uint sent=Native.Paste();
                int delay=200;string delayFile=Path.Combine(folder,"restore-delay-ms.txt");
                if(File.Exists(delayFile)){int configured;if(Int32.TryParse(File.ReadAllText(delayFile).Trim(),out configured))delay=OperationState.Delay(configured);}
                await Task.Delay(delay);
                string outcome=clipboard.Restore();
                string message=(sent==4?"已提交；":"提交结果不确定；")+outcome;
                if(outcome=="剪贴板已恢复")SetStatus(message);else Alert(message);
            }
            catch(InvalidOperationException error)
            {
                // Only our own operation errors are surfaced, not external COM/source payloads.
                var origin=error.TargetSite!=null?error.TargetSite.DeclaringType:null;
                var assembly=origin!=null?origin.Assembly:null;
                if(assembly==typeof(Addin).Assembly||assembly==typeof(MarkdownConverter).Assembly)Alert(error.Message);
                else Alert("操作未支持，已跳过");
            }
            catch(Exception){Alert(submitted?"提交后异常，不自动撤销或重试":"操作失败，未重试");}
            finally
            {
                if(clipboard!=null)
                {
                    if(clipboard.Published)
                    {retained=clipboard;state.Fault();Unhook();try{Persist(false);}catch(Exception){}Alert(status+"；保留快照，需手动恢复");}
                    else clipboard.Dispose();
                }
                Selection.Release(app);
                if(!parserOutstanding){state.End();if(retained==null)try{if(File.Exists(pending))File.Delete(pending);}catch(Exception){}}
                if(shuttingDown&&!state.Busy&&context!=null)context.ExitThread();
            }
        }
        private void RegisterLateCompletion(Task<ConversionResult> conversion)
        {
            conversion.ContinueWith(task=>
            {
                // End the ticket even after Shutdown stops accepting UI work.
                // The STA stop timer will observe idle and exit its own loop.
                if(task.IsFaulted){var observed=task.Exception;}
                state.End();Post(()=>Alert("超时任务已结束，需手动启用"));
            });
        }
        public void OnDisconnection(ext_DisconnectMode mode,ref Array custom){Shutdown();ForceExitSurrogate(mode);}
        public void OnBeginShutdown(ref Array custom){Shutdown();}
        public void OnStartupComplete(ref Array custom){}
        public void OnAddInsUpdate(ref Array custom){}
        private void Shutdown()
        {
            // Publish cancellation before returning to the host, not in a queued
            // callback that could be lost while the STA is starting/stopping.
            shuttingDown=true;state.Disable();
            try{var target=dispatcher;if(target!=null&&!target.IsDisposed)target.BeginInvoke((Action)delegate{if(toast!=null){try{if(!toast.IsDisposed)toast.Close();}catch(Exception){}toast=null;}Unhook();if(!state.Busy&&context!=null)context.ExitThread();});}catch(InvalidOperationException){}
        }
        private void ForceExitSurrogate(ext_DisconnectMode mode)
        {
            // Mirrors OneMore on the same DllSurrogate architecture: a live surrogate
            // with pending managed work can keep OneNote from closing, so this process
            // ends itself here. Only an in-flight operation waits (clipboard restore
            // stays possible); the host process is never touched.
            if(offlineLifecycle)return;
            var watch=System.Diagnostics.Stopwatch.StartNew();
            while(state.Busy&&watch.ElapsedMilliseconds<5000)Thread.Sleep(25);
            TraceLifecycle("surrogate-exit mode="+mode+(state.Busy?" busy":" idle"));
            // Whichever of the worker teardown or this forced exit runs first records the
            // single "worker-exited" line; the process ends immediately after, so the
            // record is deterministic instead of racing Environment.Exit(0).
            TraceExited();
            GC.Collect();
            Environment.Exit(0);
        }
        private void TraceExited(){if(Interlocked.Exchange(ref exitTraced,1)==0)TraceLifecycle("worker-exited");}
    }
}
