using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Collections.Generic;
using System.Threading;
using System.Windows.Forms;
using IDataObject = System.Runtime.InteropServices.ComTypes.IDataObject;

// Read-only OLE format/medium inventory. Does not request data or publish anything.
internal static class OleInventory
{
    [DllImport("ole32.dll",PreserveSig=false)]private static extern void OleGetClipboard(out IDataObject data);
    [DllImport("ole32.dll")]private static extern int OleInitialize(IntPtr reserved);
    [DllImport("ole32.dll")]private static extern void OleUninitialize();
    [DllImport("ole32.dll")]private static extern void ReleaseStgMedium(ref STGMEDIUM medium);
    [DllImport("kernel32.dll")]private static extern UIntPtr GlobalSize(IntPtr memory);
    [DllImport("kernel32.dll")]private static extern IntPtr GlobalLock(IntPtr memory);
    [DllImport("kernel32.dll")]private static extern bool GlobalUnlock(IntPtr memory);
    [DllImport("user32.dll")]private static extern uint GetClipboardSequenceNumber();
    [DllImport("user32.dll")]private static extern bool OpenClipboard(IntPtr owner);
    [DllImport("user32.dll")]private static extern bool CloseClipboard();
    [DllImport("user32.dll")]private static extern IntPtr GetClipboardOwner();
    [DllImport("user32.dll",CharSet=CharSet.Unicode)]private static extern int GetClipboardFormatName(uint id,StringBuilder name,int count);
    [STAThread]private static int Main(string[] args)
    {
        bool roundtrip=args.Length==1&&args[0]=="--roundtrip";
        if(args.Length!=0&&!roundtrip)return 2;
        int hr=OleInitialize(IntPtr.Zero);if(hr<0){Console.WriteLine("FAIL OLE initialization");return 1;}
        IDataObject data=null;IEnumFORMATETC enumeration=null;uint initialSequence=GetClipboardSequenceNumber();
        var snapshot=new Dictionary<uint,byte[]>();var omitted=new List<string>();var limits=new SnapshotLimits();
        try
        {
            OleGetClipboard(out data);enumeration=data.EnumFormatEtc(DATADIR.DATADIR_GET);
            var item=new FORMATETC[1];int count=0;
            var fetched=new int[1];
            while(enumeration.Next(1,item,fetched)==0)
            {
                try
                {
                    if(++count>128)throw new InvalidOperationException("format limit");
                    uint id=unchecked((ushort)item[0].cfFormat);
                    var name=new StringBuilder(256);GetClipboardFormatName(id,name,256);
                    // Format identifiers only, never data/source URLs/clipboard contents.
                    Console.WriteLine("OLE format="+id+" name="+(id>=0xC000?name.ToString():"[standard]")+" medium="+item[0].tymed+" aspect="+item[0].dwAspect+" index="+item[0].lindex);
                    bool protectionMarker=name.ToString()=="EnterpriseDataProtectionId";
                    bool supported=id==13||id==1||id==7||id==16||id==8||id==17||name.ToString()=="HTML Format"||name.ToString()=="Rich Text Format"||name.ToString()=="PNG";
                    if(roundtrip&&!supported)
                    {
                        if(name.ToString()=="OneNote 2016 Internal")omitted.Add("onenote-native");
                        else if(id==14||id==3||id==2)omitted.Add("graphics-fallback");
                        else throw new NotSupportedException("unapproved source representation; clipboard unchanged");
                    }
                    if((supported||(!roundtrip&&protectionMarker)) && (item[0].tymed&TYMED.TYMED_HGLOBAL)!=0 && item[0].ptd==IntPtr.Zero && item[0].lindex==-1 && item[0].dwAspect==DVASPECT.DVASPECT_CONTENT)
                    {
                        var request=item[0];request.tymed=TYMED.TYMED_HGLOBAL;
                        STGMEDIUM medium;data.GetData(ref request,out medium);
                        try
                        {
                            if(medium.tymed!=TYMED.TYMED_HGLOBAL)throw new InvalidOperationException("unexpected medium");
                            ulong size=GlobalSize(medium.unionmember).ToUInt64();
                            if(protectionMarker&&(medium.unionmember==IntPtr.Zero||size==0))
                            {
                                Console.WriteLine("Protection marker advertised but no materialized bytes; remains unclassified, not omitted in roundtrip");
                                continue;
                            }
                            if(size==0||size>4194304)throw new InvalidOperationException("snapshot size limit");
                            IntPtr pointer=GlobalLock(medium.unionmember);
                            if(pointer==IntPtr.Zero)throw new InvalidOperationException("snapshot lock");
                            if(!limits.Accept(id,(long)size))throw new InvalidOperationException("snapshot resource/duplicate limit");
                            var copy=new byte[(int)size];
                            try{Marshal.Copy(pointer,copy,0,copy.Length);}finally{GlobalUnlock(medium.unionmember);}
                            Console.WriteLine("Independent supported snapshot bytes="+copy.Length+"; contents omitted");
                            if(protectionMarker)
                            {
                                bool zero=true;foreach(byte value in copy)if(value!=0)zero=false;
                                Console.WriteLine("Protection marker all bytes zero: "+zero+"; contents not printed; no semantics inferred");
                            }
                            snapshot.Add(id,copy);
                        }
                        finally{ReleaseStgMedium(ref medium);}
                    }
                }
                finally{if(item[0].ptd!=IntPtr.Zero){Marshal.FreeCoTaskMem(item[0].ptd);item[0].ptd=IntPtr.Zero;}}
            }
            if(roundtrip)
            {
                bool rich=false;foreach(uint id in snapshot.Keys){var n=new StringBuilder(256);GetClipboardFormatName(id,n,256);if(n.ToString()=="HTML Format"||n.ToString()=="Rich Text Format")rich=true;}
                if(omitted.Contains("onenote-native")&&(!rich||!snapshot.ContainsKey(13)))throw new NotSupportedException("native omission requires rich and Unicode fallback");
                if(omitted.Contains("graphics-fallback")&&!snapshot.ContainsKey(8)&&!snapshot.ContainsKey(17))throw new NotSupportedException("graphics omission requires DIB fallback");
                if(snapshot.Count==0)throw new NotSupportedException("no common content");
                if(enumeration!=null&&Marshal.IsComObject(enumeration))Marshal.ReleaseComObject(enumeration);enumeration=null;
                if(data!=null&&Marshal.IsComObject(data))Marshal.ReleaseComObject(data);data=null;
                Roundtrip(snapshot,initialSequence);
                return 0;
            }
            Console.WriteLine("Read-only independent common-content snapshot complete; source sequence unchanged="+(initialSequence==GetClipboardSequenceNumber())+"; no publish/restore performed");return 0;
        }
        catch(Exception ex){Console.WriteLine("FAIL "+ex.GetType().Name+" HRESULT=0x"+ex.HResult.ToString("X8"));return 1;}
        finally
        {
            if(enumeration!=null&&Marshal.IsComObject(enumeration))Marshal.ReleaseComObject(enumeration);
            if(data!=null&&Marshal.IsComObject(data))Marshal.ReleaseComObject(data);
            foreach(var bytes in snapshot.Values)Array.Clear(bytes,0,bytes.Length);
            OleUninitialize();
        }
    }
    private static void Roundtrip(Dictionary<uint,byte[]> snapshot,uint initialSequence)
    {
        var original=ClipboardRoundtrip.Allocate(snapshot);
        var temporary=ClipboardRoundtrip.Allocate(new Dictionary<uint,byte[]>{{13,Encoding.Unicode.GetBytes("temporary-common-content-probe\0")}});
        using(var owner=new Owner())
        try
        {
            if(!OpenClipboard(owner.Handle))throw new InvalidOperationException("publish busy; unchanged");
            try
            {
                if(GetClipboardSequenceNumber()!=initialSequence)throw new InvalidOperationException("clipboard changed before publication");
                ClipboardRoundtrip.Write(temporary);
            }
            finally{CloseClipboard();}
            uint published=GetClipboardSequenceNumber();Thread.Sleep(200);
            bool opened=false;
            for(int i=0;i<5;i++){if(OpenClipboard(owner.Handle)){opened=true;break;}Thread.Sleep(20);}
            if(!opened)throw new InvalidOperationException("restore busy; current content may be temporary fixture");
            try
            {
                if(!RestorePolicy.MayRestore(owner.Handle.ToInt64(),GetClipboardOwner().ToInt64(),published,GetClipboardSequenceNumber()))
                {Console.WriteLine("SKIP restore; newer clipboard retained");return;}
                ClipboardRoundtrip.Write(original);
            }
            finally{CloseClipboard();}
            if(!OpenClipboard(owner.Handle))throw new InvalidOperationException("verify busy");
            try
            {
                foreach(var item in snapshot)
                {
                    IntPtr h=GetClipboardData(item.Key);ulong size=GlobalSize(h).ToUInt64();
                    if(size!=(ulong)item.Value.Length)throw new InvalidOperationException("restored size mismatch");
                    IntPtr pointer=GlobalLock(h);if(pointer==IntPtr.Zero)throw new InvalidOperationException("verify lock");
                    try{var bytes=new byte[item.Value.Length];Marshal.Copy(pointer,bytes,0,bytes.Length);for(int i=0;i<bytes.Length;i++)if(bytes[i]!=item.Value[i])throw new InvalidOperationException("restored bytes mismatch");Array.Clear(bytes,0,bytes.Length);}
                    finally{GlobalUnlock(h);}
                }
            }
            finally{CloseClipboard();}
            Console.WriteLine("PASS real-source common-content 200ms restoration; all saved representation bytes match; app-native/vector extras omitted by approved scope");
        }
        finally{ClipboardRoundtrip.Free(original);ClipboardRoundtrip.Free(temporary);}
    }
    [DllImport("user32.dll")]private static extern IntPtr GetClipboardData(uint format);
    private sealed class Owner:NativeWindow,IDisposable
    {internal Owner(){CreateHandle(new CreateParams{Parent=new IntPtr(-3)});}public void Dispose(){DestroyHandle();}}
}
