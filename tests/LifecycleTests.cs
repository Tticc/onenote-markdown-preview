using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using OneNoteMarkdown;
using Microsoft.Office.Core;
using Extensibility;
using OneNoteMarkdown.Addin;
using System.Runtime.InteropServices;

[ComVisible(true),ClassInterface(ClassInterfaceType.AutoDispatch)]
public sealed class RibbonLifetimeFixture:IRibbonUI
{
    public int Calls;
    public void Invalidate(){Calls++;}
    public void InvalidateControl(string id){Calls++;}
    public void InvalidateControlMso(string id){Calls++;}
    public void Refresh(){Calls++;}
    public void ActivateTab(string id){Calls++;}
    public void ActivateTabMso(string id){Calls++;}
    public void ActivateTabQ(string id,string ns){Calls++;}
}

// Actual Addin STA/message-loop lifecycle, with isolated config and auto-enable
// disabled. Never creates OneNote, installs a hook or touches clipboard.
internal static class LifecycleTests
{
    private static int checks;
    private static void Check(bool condition,string message)
    {if(!condition)throw new Exception(message);checks++;}
    private static FieldInfo Field(string name){return typeof(Addin).GetField(name,BindingFlags.Instance|BindingFlags.NonPublic);}
    private static Addin Create(string folder)
    {return (Addin)Activator.CreateInstance(typeof(Addin),BindingFlags.Instance|BindingFlags.NonPublic,null,new object[]{folder},null);}
    private static void Stop(Addin addin)
    {
        var thread=(Thread)Field("thread").GetValue(addin);
        Check(thread==null||thread.Join(3000),"STA did not stop");
        Check(Field("dispatcher").GetValue(addin)==null,"Dispatcher retained");
        Check(Field("context").GetValue(addin)==null,"Message context retained");
        Check((IntPtr)Field("hook").GetValue(addin)==IntPtr.Zero,"Unexpected hook");
    }
    private static void Connected(Addin addin)
    {Array custom=new object[0];addin.OnConnection(null,ext_ConnectMode.ext_cm_AfterStartup,null,ref custom);}
    private static int Main(string[] args)
    {
        if(args.Length!=1||!Directory.Exists(args[0]))return 2;
        string root=Path.Combine(args[0],"lifecycle-"+Guid.NewGuid().ToString("N"));
        try
        {
            Directory.CreateDirectory(root);
            Array custom=new object[0];
            for(int i=0;i<20;i++)
            {
                var addin=Create(Path.Combine(root,"immediate-"+i));Connected(addin);
                var watch=Stopwatch.StartNew();addin.OnDisconnection(ext_DisconnectMode.ext_dm_UserClosed,ref custom);
                Check(watch.ElapsedMilliseconds<500,"Disconnect callback waited on worker");Stop(addin);
                addin.OnBeginShutdown(ref custom); // duplicate callbacks remain harmless
            }
            int exited=0;foreach(var line in File.ReadAllLines(Path.Combine(root,"immediate-0","lifecycle.log")))if(line.Contains("worker-exited"))exited++;
            Check(exited==1,"worker-exited trace not exactly once");
            var running=Create(Path.Combine(root,"running"));Connected(running);
            var deadline=Stopwatch.StartNew();
            while(Field("dispatcher").GetValue(running)==null&&deadline.ElapsedMilliseconds<3000)Thread.Sleep(5);
            Check(Field("dispatcher").GetValue(running)!=null,"Worker not started");
            Thread original=(Thread)Field("thread").GetValue(running);Connected(running);
            Check(Object.ReferenceEquals(original,Field("thread").GetValue(running)),"Duplicate worker");
            running.OnBeginShutdown(ref custom);running.OnDisconnection(ext_DisconnectMode.ext_dm_HostShutdown,ref custom);Stop(running);
            Connected(running);Check(!((Thread)Field("thread").GetValue(running)).IsAlive,"Disconnected instance restarted");

            var busy=Create(Path.Combine(root,"busy"));Connected(busy);deadline.Restart();
            while(Field("dispatcher").GetValue(busy)==null&&deadline.ElapsedMilliseconds<3000)Thread.Sleep(5);
            var state=(OperationState)Field("state").GetValue(busy);
            Check(state.Enable(),"Synthetic state enable");long ticket=state.Begin();Check(ticket!=0,"Synthetic operation begin");
            var completion=new TaskCompletionSource<ConversionResult>();
            typeof(Addin).GetMethod("RegisterLateCompletion",BindingFlags.Instance|BindingFlags.NonPublic).Invoke(busy,new object[]{completion.Task});
            busy.OnDisconnection(ext_DisconnectMode.ext_dm_UserClosed,ref custom);
            Check(!state.Valid(ticket),"Disconnect did not invalidate pending result");
            Check(((Thread)Field("thread").GetValue(busy)).IsAlive,"Busy worker exited before cleanup finished");
            completion.SetResult(null);Stop(busy);
            Check(!state.Busy,"Late task did not complete state after shutdown");

            var ribbonFixture=new RibbonLifetimeFixture();
            var noHostBackcall=Create(Path.Combine(root,"ribbon"));
            noHostBackcall.RibbonLoaded(ribbonFixture);Connected(noHostBackcall);
            deadline.Restart();while(Field("dispatcher").GetValue(noHostBackcall)==null&&deadline.ElapsedMilliseconds<3000)Thread.Sleep(5);
            typeof(Addin).GetMethod("SetStatus",BindingFlags.Instance|BindingFlags.NonPublic).Invoke(noHostBackcall,new object[]{"fixture-only status"});
            noHostBackcall.OnDisconnection(ext_DisconnectMode.ext_dm_UserClosed,ref custom);Stop(noHostBackcall);
            Check(ribbonFixture.Calls==0,"Lifecycle/status called back into host Ribbon");
            foreach(var field in typeof(Addin).GetFields(BindingFlags.Instance|BindingFlags.NonPublic))
                Check(field.FieldType!=typeof(IRibbonUI),"Host Ribbon field retained");

            string staleFolder=Path.Combine(root,"pending");Directory.CreateDirectory(staleFolder);File.WriteAllText(Path.Combine(staleFolder,"enabled"),"1");File.WriteAllText(Path.Combine(staleFolder,"operation.pending"),"pending");
            var stale=Create(staleFolder);Connected(stale);deadline.Restart();
            while(stale.GetStatus(null).IndexOf("上次操作",StringComparison.Ordinal)<0&&deadline.ElapsedMilliseconds<3000)Thread.Sleep(5);
            Check(!((OperationState)Field("state").GetValue(stale)).Enabled,"Unfinished operation enabled automatically");
            Check(stale.GetStatus(null).Contains("上次操作"),"Pending marker not surfaced");
            stale.OnDisconnection(ext_DisconnectMode.ext_dm_UserClosed,ref custom);Stop(stale);

            string lockFolder=Path.Combine(root,"locked");Directory.CreateDirectory(lockFolder);File.WriteAllText(Path.Combine(lockFolder,"enabled"),"1");File.WriteAllText(Path.Combine(lockFolder,"lifecycle-only"),"1");
            var locked=Create(lockFolder);Connected(locked);deadline.Restart();
            while(!locked.GetStatus(null).Contains("生命周期诊断")&&deadline.ElapsedMilliseconds<3000)Thread.Sleep(5);
            Check(locked.GetStatus(null).Contains("生命周期诊断"),"Lifecycle-only guard not loaded");
            Check(!((OperationState)Field("state").GetValue(locked)).Enabled,"Lifecycle-only enabled unexpectedly");
            locked.FormatClicked(null);locked.EnableClicked(null);Thread.Sleep(100);
            Check(!((OperationState)Field("state").GetValue(locked)).Enabled,"Lifecycle-only enable click accepted");
            Check((IntPtr)Field("hook").GetValue(locked)==IntPtr.Zero,"Lifecycle-only hook installed");
            Check(!File.Exists(Path.Combine(lockFolder,"operation.pending")),"Lifecycle-only format began");
            locked.OnDisconnection(ext_DisconnectMode.ext_dm_UserClosed,ref custom);Stop(locked);

            var neverConnected=Create(Path.Combine(root,"never"));neverConnected.OnBeginShutdown(ref custom);Connected(neverConnected);
            Check(Field("thread").GetValue(neverConnected)==null,"Startup after early shutdown");
            Console.WriteLine("PASS "+checks+" actual Addin lifecycle/race/cleanup checks; no host/registration/hook/clipboard use");return 0;
        }
        catch(Exception error){Console.WriteLine("FAIL "+error.GetType().Name+" "+error.Message);return 1;}
        finally{if(Directory.Exists(root))Directory.Delete(root,true);}
    }
}
