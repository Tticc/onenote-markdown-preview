using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Windows.Forms;
using DataObject=System.Runtime.InteropServices.ComTypes.IDataObject;
namespace OneNoteMarkdown.Addin
{
    internal sealed class ClipboardSession:NativeWindow,IDisposable
    {
        [DllImport("ole32.dll",PreserveSig=false)]private static extern void OleGetClipboard(out DataObject data);
        [DllImport("ole32.dll")]private static extern void ReleaseStgMedium(ref STGMEDIUM medium);
        [DllImport("user32.dll")]private static extern bool OpenClipboard(IntPtr owner);
        [DllImport("user32.dll")]private static extern bool CloseClipboard();
        [DllImport("user32.dll")]private static extern bool EmptyClipboard();
        [DllImport("user32.dll")]private static extern IntPtr SetClipboardData(uint id,IntPtr memory);
        [DllImport("user32.dll")]private static extern IntPtr GetClipboardOwner();
        [DllImport("user32.dll")]private static extern uint GetClipboardSequenceNumber();
        [DllImport("user32.dll",CharSet=CharSet.Unicode)]private static extern uint RegisterClipboardFormat(string name);
        [DllImport("user32.dll",CharSet=CharSet.Unicode)]private static extern int GetClipboardFormatName(uint id,StringBuilder name,int length);
        [DllImport("kernel32.dll")]private static extern UIntPtr GlobalSize(IntPtr memory);
        [DllImport("kernel32.dll")]private static extern IntPtr GlobalAlloc(uint flags,UIntPtr size);
        [DllImport("kernel32.dll")]private static extern IntPtr GlobalLock(IntPtr memory);
        [DllImport("kernel32.dll")]private static extern bool GlobalUnlock(IntPtr memory);
        [DllImport("kernel32.dll")]private static extern IntPtr GlobalFree(IntPtr memory);
        private readonly Dictionary<uint,byte[]> saved=new Dictionary<uint,byte[]>();
        private Dictionary<uint,IntPtr> recovery;
        // Classified auxiliary metadata formats that carry no user content and are not
        // restored. Dropping them does not lose promised clipboard content; they only
        // help the originating app's own paste. Reviewed per class; append as encountered.
        private static readonly HashSet<string> Ignorable=new HashSet<string>(StringComparer.Ordinal)
        {"Chromium internal source RFH token","Chromium internal source URL"};
        // Human-readable labels for format names a user might hit. Kept alongside the raw
        // name so the message stays understandable without losing diagnostic precision.
        private static readonly Dictionary<string,string> Friendly=new Dictionary<string,string>(StringComparer.Ordinal)
        {
            {"Shell IDList Array","资源管理器复制的文件或文件夹列表"},
            {"Shell Object Offsets","资源管理器复制的文件或文件夹列表"},
            {"FileGroupDescriptorW","文件列表"},
            {"FileGroupDescriptor","文件列表"},
            {"FileNameW","文件名"},
            {"FileName","文件名"},
            {"Preferred DropEffect","文件操作标记"},
            {"Biff12","Excel 表格内容"},
            {"Biff8","Excel 表格内容"},
            {"Biff5","Excel 表格内容"},
            {"Biff4","Excel 表格内容"},
            {"Biff3","Excel 表格内容"},
            {"Biff2","Excel 表格内容"},
            {"Csv","CSV 表格文本"},
            {"XML Spreadsheet","Excel XML 表格"},
            {"Hyperlink","超链接信息"},
            {"Object Descriptor","嵌入对象信息"},
            {"Ole Private Data","嵌入对象信息"},
            {"WordSourceURL","Word 文档来源信息"},
            {"Microsoft Word 文档","Word 文档对象"},
            {"Art::GVML ClipFormat","图形对象信息"},
            {"DibV5","图片位图"},
            {"PNG","图片"}
        };
        private static string Describe(string name,uint id)
        {
            string friendly;
            if(Friendly.TryGetValue(name,out friendly))return friendly+"（"+name+"）";
            return name.Length==0?"未知格式 #"+id:name;
        }
        private uint originalSequence,publishedSequence;
        internal bool Published{get;private set;}
        internal bool Retained{get{return recovery!=null;}}
        internal ClipboardSession(){CreateHandle(new CreateParams{Parent=new IntPtr(-3)});}
        internal void Capture()
        {
            if(saved.Count!=0||recovery!=null)throw new InvalidOperationException("snapshot already captured");
            DataObject data=null;IEnumFORMATETC formats=null;originalSequence=GetClipboardSequenceNumber();long total=0;
            bool native=false,graphics=false,rich=false,restricted=false;
            try
            {
                OleGetClipboard(out data);formats=data.EnumFormatEtc(DATADIR.DATADIR_GET);
                var entries=new FORMATETC[1];var fetched=new int[1];int count=0;
                while(formats.Next(1,entries,fetched)==0)
                {
                    try
                    {
                        if(++count>128)throw new InvalidOperationException("剪贴板格式过多");
                        uint id=unchecked((ushort)entries[0].cfFormat);var label=new StringBuilder(256);GetClipboardFormatName(id,label,256);string name=label.ToString();
                        if(name=="OneNote 2016 Internal"){native=true;continue;}
                        if(id==2||id==3||id==14){graphics=true;continue;}
                        if(Ignorable.Contains(name))continue;
                        bool supported=id==1||id==7||id==13||id==16||id==8||id==17||name=="HTML Format"||name=="Rich Text Format"||name=="PNG";
                        bool privacy=name=="CanIncludeInClipboardHistory"||name=="CanUploadToCloudClipboard"||name=="ExcludeClipboardContentFromMonitorProcessing";
                        if(!supported&&!privacy)throw new InvalidOperationException("剪贴板含未支持的内容："+Describe(name,id)+"，已跳过");
                        var request=entries[0];
                        if(request.ptd!=IntPtr.Zero||request.lindex!=-1||request.dwAspect!=DVASPECT.DVASPECT_CONTENT||(request.tymed&TYMED.TYMED_HGLOBAL)==0)throw new InvalidOperationException("剪贴板格式存储形式未支持："+Describe(name,id));
                        request.tymed=TYMED.TYMED_HGLOBAL;STGMEDIUM medium;data.GetData(ref request,out medium);
                        try
                        {
                            ulong size=GlobalSize(medium.unionmember).ToUInt64();
                            if(medium.tymed!=TYMED.TYMED_HGLOBAL||size==0||size>4194304||total+(long)size>8388608||saved.ContainsKey(id))throw new InvalidOperationException("剪贴板格式大小或结构超限："+Describe(name,id));
                            var bytes=new byte[(int)size];IntPtr pointer=GlobalLock(medium.unionmember);if(pointer==IntPtr.Zero)throw new InvalidOperationException("无法读取剪贴板");
                            try{Marshal.Copy(pointer,bytes,0,bytes.Length);}finally{GlobalUnlock(medium.unionmember);}
                            if(privacy&&name!="ExcludeClipboardContentFromMonitorProcessing"&&(bytes.Length!=4||(BitConverter.ToUInt32(bytes,0)!=0&&BitConverter.ToUInt32(bytes,0)!=1)))throw new InvalidOperationException("隐私标记格式未支持");
                            if(privacy)restricted=true;
                            saved.Add(id,bytes);total+=bytes.Length;
                            if(name=="HTML Format"||name=="Rich Text Format")rich=true;
                        }
                        finally{ReleaseStgMedium(ref medium);}
                    }
                    finally{if(entries[0].ptd!=IntPtr.Zero){Marshal.FreeCoTaskMem(entries[0].ptd);entries[0].ptd=IntPtr.Zero;}}
                }
                if(native&&(!rich||!saved.ContainsKey(13)))throw new InvalidOperationException("OneNote 专属剪贴板数据缺少富文本后备，已跳过");
                if(graphics&&!saved.ContainsKey(8)&&!saved.ContainsKey(17))throw new InvalidOperationException("剪贴板含不支持的图片格式（缺少位图后备），已跳过");
                if(restricted)throw new InvalidOperationException("剪贴板带隐私保护标记，暂不支持，已跳过");
                if(GetClipboardSequenceNumber()!=originalSequence)throw new InvalidOperationException("读取时剪贴板已变化");
                recovery=Allocate(saved);
            }
            finally{Selection.Release(formats);Selection.Release(data);}
        }
        internal void Publish(ConversionResult converted)
        {
            if(recovery==null)throw new InvalidOperationException("missing snapshot");
            var data=new Dictionary<uint,byte[]>();
            data.Add(RegisterClipboardFormat("HTML Format"),HtmlClipboard.Create(converted.Html));
            data.Add(13,Encoding.Unicode.GetBytes(converted.Text+"\0"));
            // Temporary notes must not pollute Windows history/cloud by default.
            data.Add(RegisterClipboardFormat("CanIncludeInClipboardHistory"),BitConverter.GetBytes(0));
            data.Add(RegisterClipboardFormat("CanUploadToCloudClipboard"),BitConverter.GetBytes(0));
            var handles=Allocate(data);
            try
            {
                if(!OpenClipboard(Handle))throw new InvalidOperationException("剪贴板被占用");
                try
                {
                    if(GetClipboardSequenceNumber()!=originalSequence)throw new InvalidOperationException("剪贴板已变化");
                    Published=true;Write(handles);
                }
                finally{CloseClipboard();if(Published)publishedSequence=GetClipboardSequenceNumber();}
            }
            finally{Free(handles);foreach(var bytes in data.Values)Array.Clear(bytes,0,bytes.Length);}
        }
        internal string Restore()
        {
            if(!Published)return "未发布剪贴板";
            if(!OpenClipboard(Handle))return "剪贴板忙，快照待恢复";
            bool terminal=false;string status;
            try
            {
                if(GetClipboardOwner()!=Handle||GetClipboardSequenceNumber()!=publishedSequence){terminal=true;status="保留后来复制的内容";}
                else
                {
                    // Reallocate before EmptyClipboard: original independent bytes survive partial failures.
                    var handles=Allocate(saved);
                    try{Write(handles);terminal=true;status="剪贴板已恢复";}
                    finally{Free(handles);}
                }
            }
            finally{CloseClipboard();}
            if(terminal){Published=false;Free(recovery);recovery=null;}
            return status;
        }
        private static Dictionary<uint,IntPtr> Allocate(Dictionary<uint,byte[]> bytes)
        {
            var handles=new Dictionary<uint,IntPtr>();
            try
            {
                foreach(var item in bytes)
                {
                    IntPtr memory=GlobalAlloc(2,new UIntPtr((uint)item.Value.Length));if(memory==IntPtr.Zero)throw new OutOfMemoryException();handles.Add(item.Key,memory);
                    IntPtr pointer=GlobalLock(memory);if(pointer==IntPtr.Zero)throw new InvalidOperationException("分配快照失败");
                    try{Marshal.Copy(item.Value,0,pointer,item.Value.Length);}finally{GlobalUnlock(memory);}
                }
                return handles;
            }
            catch{Free(handles);throw;}
        }
        private static void Write(Dictionary<uint,IntPtr> handles)
        {
            if(!EmptyClipboard())throw new InvalidOperationException("清空剪贴板失败");
            foreach(uint id in new List<uint>(handles.Keys))
            {if(SetClipboardData(id,handles[id])==IntPtr.Zero)throw new InvalidOperationException("发布剪贴板失败");handles[id]=IntPtr.Zero;}
        }
        private static void Free(Dictionary<uint,IntPtr> handles){if(handles!=null)foreach(var handle in handles.Values)if(handle!=IntPtr.Zero)GlobalFree(handle);}
        public void Dispose(){Free(recovery);recovery=null;foreach(var bytes in saved.Values)Array.Clear(bytes,0,bytes.Length);saved.Clear();DestroyHandle();}
    }
}
