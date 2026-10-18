using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Diagnostics;
using System.Windows.Forms;

// Authorized synthetic fixtures only. Leaves the last image fixture on clipboard.
internal static class CommonClipboardTests
{
    [DllImport("user32.dll")] private static extern bool OpenClipboard(IntPtr w);
    [DllImport("user32.dll")] private static extern bool CloseClipboard();
    [DllImport("user32.dll")] private static extern uint GetClipboardSequenceNumber();
    [DllImport("user32.dll")] private static extern IntPtr GetClipboardOwner();
    [DllImport("user32.dll",SetLastError=true)] private static extern uint EnumClipboardFormats(uint f);
    [DllImport("user32.dll")] private static extern IntPtr GetClipboardData(uint f);
    [DllImport("user32.dll",CharSet=CharSet.Unicode)] private static extern uint RegisterClipboardFormat(string name);
    [DllImport("kernel32.dll")] private static extern UIntPtr GlobalSize(IntPtr h);
    [DllImport("kernel32.dll")] private static extern IntPtr GlobalLock(IntPtr h);
    [DllImport("kernel32.dll")] private static extern bool GlobalUnlock(IntPtr h);
    [DllImport("kernel32.dll")] private static extern void SetLastError(uint e);
    private static uint html,rtf,png;
    private static void Lock(IntPtr owner)
    {
        for(int i=0;i<10;i++){if(OpenClipboard(owner))return;Thread.Sleep(20);}
        throw new InvalidOperationException("clipboard busy after bounded attempts");
    }
    private static void Publish(IntPtr owner,Dictionary<uint,byte[]> data)
    {
        var handles=ClipboardRoundtrip.Allocate(data);
        try{Lock(owner);try{ClipboardRoundtrip.Write(handles);}finally{CloseClipboard();}}
        finally{ClipboardRoundtrip.Free(handles);}
    }
    private static Dictionary<uint,byte[]> Snapshot(IntPtr owner)
    {
        Lock(owner);
        try
        {
            var formats=new List<uint>();uint format=0;
            while(true)
            {
                SetLastError(0);format=EnumClipboardFormats(format);
                if(format==0){if(Marshal.GetLastWin32Error()!=0)throw new InvalidOperationException("enumeration");break;}
                formats.Add(format);if(formats.Count>32)throw new InvalidOperationException("format limit");
            }
            var snapshot=new Dictionary<uint,byte[]>();
            foreach(uint f in formats)
            {
                if(f==2 && (formats.Contains(8)||formats.Contains(17)))continue; // synthesized bitmap; preserve DIB representations
                if(f!=1&&f!=7&&f!=13&&f!=16&&f!=8&&f!=17&&f!=html&&f!=rtf&&f!=png)throw new NotSupportedException("unapproved format");
                IntPtr handle=GetClipboardData(f);ulong size=GlobalSize(handle).ToUInt64();
                if(handle==IntPtr.Zero||size==0||size>1048576)throw new InvalidOperationException("medium limit");
                IntPtr pointer=GlobalLock(handle);if(pointer==IntPtr.Zero)throw new InvalidOperationException("medium lock");
                try{var bytes=new byte[(int)size];Marshal.Copy(pointer,bytes,0,bytes.Length);snapshot.Add(f,bytes);}
                finally{GlobalUnlock(handle);}
            }
            return snapshot;
        }
        finally{CloseClipboard();}
    }
    private static bool Same(byte[] a,byte[] b)
    {if(a.Length!=b.Length)return false;for(int i=0;i<a.Length;i++)if(a[i]!=b[i])return false;return true;}
    private static void Roundtrip(IntPtr owner,Dictionary<uint,byte[]> fixture,string label)
    {
        Publish(owner,fixture);var snapshot=Snapshot(owner);
        foreach(var expected in fixture)if(!snapshot.ContainsKey(expected.Key)||!Same(expected.Value,snapshot[expected.Key]))throw new InvalidOperationException("fixture altered before snapshot");
        var restore=ClipboardRoundtrip.Allocate(snapshot);
        try
        {
            Publish(owner,new Dictionary<uint,byte[]>{{13,Encoding.Unicode.GetBytes("temporary\0")}});
            uint sequence=GetClipboardSequenceNumber();Thread.Sleep(200);
            Lock(owner);
            try
            {
                if(!RestorePolicy.MayRestore(owner.ToInt64(),GetClipboardOwner().ToInt64(),sequence,GetClipboardSequenceNumber()))throw new InvalidOperationException("external change preserved; roundtrip canceled");
                ClipboardRoundtrip.Write(restore);
            }
            finally{CloseClipboard();}
            var actual=Snapshot(owner);
            foreach(var item in snapshot)if(!actual.ContainsKey(item.Key)||!Same(item.Value,actual[item.Key]))throw new InvalidOperationException("restored bytes differ");
            Console.WriteLine("PASS synthetic "+label+" independent snapshot/200ms restore; supported representation bytes match");
        }
        finally{ClipboardRoundtrip.Free(restore);}
    }
    private static byte[] Dib()
    {
        var bytes=new byte[56]; // BITMAPINFOHEADER + 2x2 32-bit BGRA pixels
        Array.Copy(BitConverter.GetBytes(40),0,bytes,0,4);
        Array.Copy(BitConverter.GetBytes(2),0,bytes,4,4);Array.Copy(BitConverter.GetBytes(2),0,bytes,8,4);
        Array.Copy(BitConverter.GetBytes((short)1),0,bytes,12,2);Array.Copy(BitConverter.GetBytes((short)32),0,bytes,14,2);
        Array.Copy(BitConverter.GetBytes(16),0,bytes,20,4);
        for(int i=0;i<4;i++){bytes[40+i*4]=(byte)(i*40);bytes[41+i*4]=100;bytes[42+i*4]=200;bytes[43+i*4]=255;}
        return bytes;
    }
    [STAThread] private static int Main(string[] args)
    {
        using(var owner=new Owner())
        try
        {
            if(args.Length==1 && args[0]=="--competing-copy")
            {
                Publish(owner.Handle,new Dictionary<uint,byte[]>{{13,Encoding.Unicode.GetBytes("external-process-fixture\0")}});
                return 0;
            }
            if(args.Length!=0)return 2;
            html=RegisterClipboardFormat("HTML Format");rtf=RegisterClipboardFormat("Rich Text Format");png=RegisterClipboardFormat("PNG");
            Roundtrip(owner.Handle,new Dictionary<uint,byte[]>{{13,Encoding.Unicode.GetBytes("中文😀\0")}},"Unicode text");
            Roundtrip(owner.Handle,new Dictionary<uint,byte[]>{
                {html,HtmlFixturePayload.Create("<b>中文😀</b>")},
                {rtf,Encoding.ASCII.GetBytes("{\\rtf1\\ansi\\uc1 \\b test\\b0 \\u20013?\\u25991?}\0")},
                {13,Encoding.Unicode.GetBytes("中文😀\0")}},"HTML/RTF/text");
            Publish(owner.Handle,new Dictionary<uint,byte[]>{{13,Encoding.Unicode.GetBytes("original-disposable\0")}});
            uint prior=GetClipboardSequenceNumber();
            using(var other=Process.Start(new ProcessStartInfo(Process.GetCurrentProcess().MainModule.FileName,"--competing-copy"){UseShellExecute=false}))
            {
                if(!other.WaitForExit(3000)){other.Kill();throw new InvalidOperationException("competing fixture timed out");}
                if(other.ExitCode!=0)throw new InvalidOperationException("competing fixture failed");
            }
            Lock(owner.Handle);
            try
            {
                if(RestorePolicy.MayRestore(owner.Handle.ToInt64(),GetClipboardOwner().ToInt64(),prior,GetClipboardSequenceNumber()))throw new InvalidOperationException("competing process would be overwritten");
                var handle=GetClipboardData(13);var pointer=GlobalLock(handle);
                if(pointer==IntPtr.Zero)throw new InvalidOperationException("newer text unavailable");
                try{if(Marshal.PtrToStringUni(pointer)!="external-process-fixture")throw new InvalidOperationException("newer content lost");}
                finally{GlobalUnlock(handle);}
            }
            finally{CloseClipboard();}
            Console.WriteLine("PASS competing-process copy rejected by restore guard; newer text retained");
            byte[] image;
            using(var bitmap=new Bitmap(2,2))using(var stream=new System.IO.MemoryStream())
            {
                bitmap.SetPixel(0,0,Color.Red);bitmap.SetPixel(1,0,Color.Green);bitmap.SetPixel(0,1,Color.Blue);bitmap.SetPixel(1,1,Color.White);
                bitmap.Save(stream,ImageFormat.Png);image=stream.ToArray();
            }
            Roundtrip(owner.Handle,new Dictionary<uint,byte[]>{{png,image}},"PNG stream");
            Roundtrip(owner.Handle,new Dictionary<uint,byte[]>{{8,Dib()}},"2x2 DIB (synthesized bitmap omitted)");
            using(var restored=Clipboard.GetImage())
            {
                if(restored==null||restored.Width!=2||restored.Height!=2)throw new InvalidOperationException("image dimensions");
                using(var bitmap=new Bitmap(restored))
                    for(int y=0;y<2;y++)for(int x=0;x<2;x++)
                    {var pixel=bitmap.GetPixel(x,y);int i=(1-y)*2+x;if(pixel.R!=200||pixel.G!=100||pixel.B!=i*40)throw new InvalidOperationException("pixel mismatch");}
            }
            Console.WriteLine("PASS restored DIB decoded dimensions/pixels; last disposable image remains; no OneNote access");return 0;
        }
        catch(Exception ex){Console.WriteLine("FAIL "+ex.GetType().Name+": "+ex.Message+"; clipboard may retain a disposable fixture");return 1;}
    }
    private sealed class Owner:NativeWindow,IDisposable
    {internal Owner(){CreateHandle(new CreateParams{Parent=new IntPtr(-3)});}public void Dispose(){DestroyHandle();}}
}
