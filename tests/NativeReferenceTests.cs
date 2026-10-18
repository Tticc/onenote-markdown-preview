using System;
using System.IO;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Threading;
using Extensibility;
using OneNoteMarkdown.Addin;

// Real unmanaged IDispatch memory/vtable with measured AddRef/Release, not a
// managed mock posing as an RCW. Never activates/registers COM or touches Office.
internal static class NativeReferenceTests
{
    [UnmanagedFunctionPointer(CallingConvention.StdCall)]private delegate int Query(IntPtr self,ref Guid iid,out IntPtr value);
    [UnmanagedFunctionPointer(CallingConvention.StdCall)]private delegate uint Ref(IntPtr self);
    [UnmanagedFunctionPointer(CallingConvention.StdCall)]private delegate int TypeCount(IntPtr self,out uint count);
    [UnmanagedFunctionPointer(CallingConvention.StdCall)]private delegate int TypeInfo(IntPtr self,uint index,uint locale,out IntPtr result);
    [UnmanagedFunctionPointer(CallingConvention.StdCall)]private delegate int Names(IntPtr self,IntPtr iid,IntPtr names,uint count,uint locale,IntPtr ids);
    [UnmanagedFunctionPointer(CallingConvention.StdCall)]private delegate int Invoke(IntPtr self,int id,IntPtr iid,uint locale,ushort flags,IntPtr args,IntPtr result,IntPtr error,IntPtr argError);
    private sealed class NativeObject:IDisposable
    {
        internal IntPtr Pointer;private IntPtr table;internal int References=1,Adds,Releases;
        private readonly Delegate[] delegates;
        internal NativeObject()
        {
            delegates=new Delegate[]{new Query(QI),new Ref(Add),new Ref(Release),new TypeCount(Count),new TypeInfo(Info),new Names(GetNames),new Invoke(Call)};
            table=Marshal.AllocCoTaskMem(IntPtr.Size*delegates.Length);Pointer=Marshal.AllocCoTaskMem(IntPtr.Size);
            for(int i=0;i<delegates.Length;i++)Marshal.WriteIntPtr(table,i*IntPtr.Size,Marshal.GetFunctionPointerForDelegate(delegates[i]));
            Marshal.WriteIntPtr(Pointer,table);
        }
        private int QI(IntPtr self,ref Guid iid,out IntPtr value)
        {
            if(iid==new Guid("00000000-0000-0000-C000-000000000046")||iid==new Guid("00020400-0000-0000-C000-000000000046"))
            {Add(self);value=Pointer;return 0;}
            value=IntPtr.Zero;return unchecked((int)0x80004002);
        }
        private uint Add(IntPtr self){Interlocked.Increment(ref Adds);return (uint)Interlocked.Increment(ref References);}
        private uint Release(IntPtr self){Interlocked.Increment(ref Releases);return (uint)Interlocked.Decrement(ref References);}
        private int Count(IntPtr self,out uint count){count=0;return 0;}
        private int Info(IntPtr self,uint index,uint locale,out IntPtr result){result=IntPtr.Zero;return unchecked((int)0x80004001);}
        private int GetNames(IntPtr self,IntPtr iid,IntPtr names,uint count,uint locale,IntPtr ids){return unchecked((int)0x80020006);}
        private int Call(IntPtr self,int id,IntPtr iid,uint locale,ushort flags,IntPtr args,IntPtr result,IntPtr error,IntPtr argError){return unchecked((int)0x80020003);}
        public void Dispose(){if(References!=1)throw new InvalidOperationException("Native object still referenced: "+References);Marshal.FreeCoTaskMem(Pointer);Marshal.FreeCoTaskMem(table);GC.KeepAlive(delegates);}
    }
    [MethodImpl(MethodImplOptions.NoInlining)]
    private static WeakReference Connect(Addin addin,NativeObject application,NativeObject owner)
    {
        object a=Marshal.GetObjectForIUnknown(application.Pointer),b=Marshal.GetObjectForIUnknown(owner.Pointer);
        if(!Marshal.IsComObject(a)||!Marshal.IsComObject(b))throw new Exception("Fixture is not real RCW");
        Array custom=new object[0];addin.OnConnection(a,ext_ConnectMode.ext_cm_AfterStartup,b,ref custom);
        return new WeakReference(a);
    }
    [STAThread]private static int Main(string[] args)
    {
        if(args.Length!=1||!Directory.Exists(args[0]))return 2;
        string folder=Path.Combine(args[0],"native-ref-"+Guid.NewGuid().ToString("N"));
        // Memory is freed only after finalizers release all measured native refs.
        var application=new NativeObject();var owner=new NativeObject();
        try
        {
            var addin=(Addin)Activator.CreateInstance(typeof(Addin),BindingFlags.Instance|BindingFlags.NonPublic,null,new object[]{folder},null);
            var weak=Connect(addin,application,owner);
            Console.WriteLine("OBSERVE returned from callback: native application refs="+application.References+" addin refs="+owner.References);
            if(application.References<=1||owner.References<=1)throw new Exception("Expected real RCW ownership was not observed");
            Array custom=new object[0];addin.OnDisconnection(ext_DisconnectMode.ext_dm_UserClosed,ref custom);
            var thread=(Thread)typeof(Addin).GetField("thread",BindingFlags.Instance|BindingFlags.NonPublic).GetValue(addin);
            if(thread!=null&&!thread.Join(3000))throw new Exception("Worker did not end");
            Console.WriteLine("OBSERVE worker ended: native application refs="+application.References+" addin refs="+owner.References);
            for(int i=0;i<5&&(application.References!=1||owner.References!=1);i++)
            {GC.Collect();GC.WaitForPendingFinalizers();GC.Collect();Thread.Sleep(10);}
            if(application.References!=1||owner.References!=1)throw new Exception("Native RCW references remain after collection");
            Console.WriteLine("PASS native host argument refs return to fixture baseline after GC; weak alive="+weak.IsAlive+" application add/release="+application.Adds+"/"+application.Releases);
            Console.WriteLine("This demonstrates RCW release timing, not the OneNote incident root cause or cross-process RPC rundown");return 0;
        }
        catch(Exception ex){Console.WriteLine("FAIL "+ex.GetType().Name+" "+ex.Message);return 1;}
        finally
        {
            if(application.References==1)application.Dispose();
            if(owner.References==1)owner.Dispose();
            if(Directory.Exists(folder))Directory.Delete(folder,true);
        }
    }
}
