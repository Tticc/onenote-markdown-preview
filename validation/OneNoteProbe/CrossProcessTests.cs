using System;
using System.Diagnostics;
using System.IO;
using System.IO.Pipes;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Security.AccessControl;
using System.Security.Principal;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Extensibility;
using Microsoft.Office.Core;
using DISPPARAMS = System.Runtime.InteropServices.ComTypes.DISPPARAMS;
using EXCEPINFO = System.Runtime.InteropServices.ComTypes.EXCEPINFO;
#if PRODUCT_TEST
using ProbeAddin = OneNoteMarkdown.Addin.Addin;
#endif

[ComVisible(true), ClassInterface(ClassInterfaceType.AutoDispatch)]
public sealed class TestHostObject
{
    public string Name { get { return "Offline fixture"; } }
}

[ComVisible(true), ClassInterface(ClassInterfaceType.AutoDispatch)]
public sealed class TestRibbonControl : IRibbonControl
{
    public int IdReads;
    public string Id { get { IdReads++; return "OfflineButton"; } }
    public object Context { get { return null; } }
    public string Tag { get { return ""; } }
}

[ComImport, Guid("00020400-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface NativeDispatch
{
    [PreserveSig] int GetTypeInfoCount(out uint count);
    [PreserveSig] int GetTypeInfo(uint index, uint locale, out IntPtr info);
    [PreserveSig] int GetIDsOfNames(ref Guid iid, [MarshalAs(UnmanagedType.LPArray, ArraySubType=UnmanagedType.LPWStr, SizeParamIndex=2)] string[] names, uint count, uint locale, [Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex=2)] int[] ids);
    [PreserveSig] int Invoke(int id, ref Guid iid, uint locale, ushort flags, ref DISPPARAMS args, IntPtr result, ref EXCEPINFO error, out uint argumentError);
}

// Explicit COM marshaling over a private test pipe. No registry or OneNote activation.
internal static class CrossProcessTests
{
    [DllImport("ole32.dll", PreserveSig = false)]
    private static extern void CreateStreamOnHGlobal(IntPtr memory, bool deleteOnRelease, out IStream stream);
    [DllImport("ole32.dll", PreserveSig = false)]
    private static extern void CoMarshalInterface(IStream stream, ref Guid iid, IntPtr instance, uint context, IntPtr contextData, uint flags);
    [DllImport("ole32.dll", PreserveSig = false)]
    private static extern void CoUnmarshalInterface(IStream stream, ref Guid iid, out IntPtr instance);
    [DllImport("oleaut32.dll")] private static extern int VariantClear(IntPtr variant);

    [STAThread]
    private static int Main(string[] args)
    {
        try { return args.Length == 2 && args[0] == "--server" ? Server(args[1]) : Client(); }
        catch (Exception ex)
        {
            Console.WriteLine("FAIL " + ex.GetType().Name + " HRESULT=0x" + ex.HResult.ToString("X8"));
            return 1;
        }
    }

    private static int Server(string name)
    {
        var security = new PipeSecurity();
        security.SetAccessRuleProtection(true, false);
        security.AddAccessRule(new PipeAccessRule(WindowsIdentity.GetCurrent().User, PipeAccessRights.FullControl, AccessControlType.Allow));
        using (var pipe = new NamedPipeServerStream(name, PipeDirection.InOut, 1, PipeTransmissionMode.Byte, PipeOptions.None, 4096, 4096, security))
        {
            // Parent watchdog owns the bounded lifetime of this test process.
            pipe.WaitForConnection();
#if PRODUCT_TEST
            string fixtureFolder=Path.Combine(Path.GetTempPath(),"OneNoteMarkdownLifecycle-"+Guid.NewGuid().ToString("N"));
            var probe=(ProbeAddin)Activator.CreateInstance(typeof(ProbeAddin),System.Reflection.BindingFlags.Instance|System.Reflection.BindingFlags.NonPublic,null,new object[]{fixtureFolder},null);
#else
            var probe = new ProbeAddin();
#endif
            IStream stream;
            CreateStreamOnHGlobal(IntPtr.Zero, true, out stream);
            IntPtr unknown = Marshal.GetIDispatchForObject(probe);
            try
            {
                Guid iid = typeof(NativeDispatch).GUID;
                CoMarshalInterface(stream, ref iid, unknown, 0, IntPtr.Zero, 0);
                System.Runtime.InteropServices.ComTypes.STATSTG stat;
                stream.Stat(out stat, 1);
                if (stat.cbSize <= 0 || stat.cbSize > 65536) throw new InvalidDataException();
                byte[] data = new byte[(int)stat.cbSize];
                stream.Seek(0, 0, IntPtr.Zero);
                stream.Read(data, data.Length, IntPtr.Zero);
                var writer = new BinaryWriter(pipe);
                writer.Write(Process.GetCurrentProcess().Id);
                writer.Write(data.Length);
                writer.Write(data);
                writer.Flush();
                var stop = Task.Run(() => pipe.ReadByte());
                while (!stop.IsCompleted) { Application.DoEvents(); Thread.Sleep(5); }
                if (stop.Result != 1) return 2;
#if PRODUCT_TEST
                var worker=(Thread)typeof(ProbeAddin).GetField("thread",System.Reflection.BindingFlags.Instance|System.Reflection.BindingFlags.NonPublic).GetValue(probe);
                var deadline=Stopwatch.StartNew();
                while(worker!=null&&worker.IsAlive&&deadline.ElapsedMilliseconds<3000){Application.DoEvents();Thread.Sleep(5);}
                if(worker!=null&&worker.IsAlive)throw new InvalidOperationException("Product STA did not exit");
                if(Directory.Exists(fixtureFolder))Directory.Delete(fixtureFolder,true);
                Console.WriteLine("PASS actual product STA stopped after cross-process host disconnect");
#endif
                GC.KeepAlive(probe);
                Console.WriteLine("Server acknowledged explicit stop and is exiting normally");
                return 0;
            }
            finally { Marshal.Release(unknown); Marshal.ReleaseComObject(stream); }
        }
    }

    private static int Client()
    {
        string name = "OneNoteMarkdownOffline-" + Guid.NewGuid().ToString("N");
        using (var child = Process.Start(new ProcessStartInfo(Process.GetCurrentProcess().MainModule.FileName, "--server " + name) { UseShellExecute = false }))
        using (var watchdog = new System.Threading.Timer(delegate {
            try { if (!child.HasExited) child.Kill(); } catch { }
            Console.WriteLine("FAIL cross-process test exceeded 20 seconds");
            Environment.Exit(3);
        }, null, 20000, Timeout.Infinite))
        {
            try
            {
                using (var pipe = new NamedPipeClientStream(".", name, PipeDirection.InOut))
                {
                    pipe.Connect(5000);
                    var reader = new BinaryReader(pipe);
                    int serverPid = reader.ReadInt32();
                    if (serverPid != child.Id) throw new InvalidDataException();
                    int length = reader.ReadInt32();
                    if (length <= 0 || length > 65536) throw new InvalidDataException();
                    byte[] data = reader.ReadBytes(length);
                    if (data.Length != length) throw new EndOfStreamException();
                    IStream stream;
                    CreateStreamOnHGlobal(IntPtr.Zero, true, out stream);
                    object remote = null;
                    try
                    {
                        stream.Write(data, data.Length, IntPtr.Zero);
                        stream.Seek(0, 0, IntPtr.Zero);
                        Guid iid = typeof(NativeDispatch).GUID;
                        IntPtr ptr;
                        CoUnmarshalInterface(stream, ref iid, out ptr);
                        try { remote = Marshal.GetObjectForIUnknown(ptr); }
                        finally { Marshal.Release(ptr); }
                        Console.WriteLine("PASS default IDispatch unmarshaled; client=" + Process.GetCurrentProcess().Id + " server=" + serverPid);
                        var lifecycle = (IDTExtensibility2)remote;
                        Array custom = new object[] { "offline-test", 42 };
                        var application = new TestHostObject();
                        var addin = new TestHostObject();
                        lifecycle.OnConnection(application, ext_ConnectMode.ext_cm_Startup, addin, ref custom);
                        lifecycle.OnAddInsUpdate(ref custom);
                        lifecycle.OnStartupComplete(ref custom);
                        Console.WriteLine("PASS cross-process lifecycle startup with SAFEARRAY");
                        string xml = ((IRibbonExtensibility)remote).GetCustomUI("offline-cross-process");
#if PRODUCT_TEST
                        if (!xml.Contains("SafeMdTab")) throw new InvalidDataException();
                        var parsed=new System.Xml.XmlDocument{XmlResolver=null};parsed.LoadXml(xml);
                        if(parsed.DocumentElement.HasAttribute("onLoad"))throw new InvalidOperationException("Host Ribbon retention unexpectedly restored");
#else
                        if (!xml.Contains("Markdown Probe")) throw new InvalidDataException();
#endif
                        Console.WriteLine("PASS cross-process official Ribbon interface and BSTR result");
#if !PRODUCT_TEST
                        var control = new TestRibbonControl();
                        var dispatch = (NativeDispatch)remote;
                        Guid empty = Guid.Empty;
                        int[] ids = new int[1];
                        Console.WriteLine("TRACE resolving VerifyCallback");
                        Marshal.ThrowExceptionForHR(dispatch.GetIDsOfNames(ref empty, new[] { "VerifyCallback" }, 1, 0, ids));
                        Console.WriteLine("TRACE callback DISPID=" + ids[0]);
                        if (ids[0] == 0 || ids[0] == -1) throw new InvalidOperationException("Invalid callback DISPID");
                        int[] missing = new int[1];
                        int missingResult = dispatch.GetIDsOfNames(ref empty, new[] { "MissingCallbackForNegativeTest" }, 1, 0, missing);
                        if (missingResult != unchecked((int)0x80020006) || missing[0] != -1)
                            throw new InvalidOperationException("Unknown-name marshaling regression");
                        Console.WriteLine("PASS unknown callback returns DISP_E_UNKNOWNNAME and DISPID_UNKNOWN");
                        IntPtr variant = Marshal.AllocCoTaskMem(32);
                        try
                        {
                            for (int i = 0; i < 32; i++) Marshal.WriteByte(variant, i, 0);
                            Marshal.GetNativeVariantForObject(control, variant);
                            var parameters = new DISPPARAMS { cArgs = 1, rgvarg = variant };
                            var error = new EXCEPINFO();
                            uint argumentError;
                            Console.WriteLine("TRACE invoking callback args=" + parameters.cArgs + " variantType=" + Marshal.ReadInt16(variant));
                            Marshal.ThrowExceptionForHR(dispatch.Invoke(ids[0], ref empty, 0, 1, ref parameters, IntPtr.Zero, ref error, out argumentError));
                            if (control.IdReads != 1) throw new InvalidOperationException("Expected reverse COM control getter call");
                            Console.WriteLine("PASS native IDispatch GetIDsOfNames/Invoke; reverse IRibbonControl.Id getter called once");
                            var emptyParameters = new DISPPARAMS();
                            var negativeError = new EXCEPINFO();
                            int badCount = dispatch.Invoke(ids[0], ref empty, 0, 1, ref emptyParameters, IntPtr.Zero, ref negativeError, out argumentError);
                            Console.WriteLine("TRACE missing argument HRESULT=0x" + badCount.ToString("X8") + " scode=0x" + negativeError.scode.ToString("X8"));
                            if (badCount >= 0 || control.IdReads != 1)
                                throw new InvalidOperationException("Bad argument count regression");
                            Console.WriteLine("PASS missing argument rejected with failing HRESULT without callback");
                        }
                        finally { VariantClear(variant); Marshal.FreeCoTaskMem(variant); }
#endif
                        lifecycle.OnBeginShutdown(ref custom);
                        lifecycle.OnDisconnection(ext_DisconnectMode.ext_dm_HostShutdown, ref custom);
                        Console.WriteLine("PASS cross-process lifecycle shutdown");
                        GC.KeepAlive(application);
                        GC.KeepAlive(addin);
#if !PRODUCT_TEST
                        GC.KeepAlive(control);
#endif
                    }
                    finally
                    {
                        if (remote != null && Marshal.IsComObject(remote)) Marshal.FinalReleaseComObject(remote);
                        Marshal.ReleaseComObject(stream);
                    }
                    pipe.WriteByte(1);
                    pipe.Flush();
                    if (!child.WaitForExit(3000) || child.ExitCode != 0) throw new InvalidOperationException();
                    Console.WriteLine("PASS client released COM object; test server exited normally");
                    return 0;
                }
            }
            finally { if (!child.HasExited) { child.Kill(); child.WaitForExit(3000); } }
        }
    }
}
