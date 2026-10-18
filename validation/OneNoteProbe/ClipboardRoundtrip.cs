using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

// Disposable plain-text fixture only. No OneNote access. Strict format gate.
internal static class ClipboardRoundtrip
{
    [DllImport("user32.dll")] private static extern bool OpenClipboard(IntPtr owner);
    [DllImport("user32.dll")] private static extern bool CloseClipboard();
    [DllImport("user32.dll")] private static extern bool EmptyClipboard();
    [DllImport("user32.dll", SetLastError=true)] private static extern uint EnumClipboardFormats(uint previous);
    [DllImport("user32.dll")] private static extern IntPtr GetClipboardData(uint format);
    [DllImport("user32.dll")] private static extern IntPtr SetClipboardData(uint format, IntPtr memory);
    [DllImport("user32.dll")] private static extern uint GetClipboardSequenceNumber();
    [DllImport("user32.dll")] private static extern IntPtr GetClipboardOwner();
    [DllImport("kernel32.dll")] private static extern UIntPtr GlobalSize(IntPtr memory);
    [DllImport("kernel32.dll")] private static extern IntPtr GlobalLock(IntPtr memory);
    [DllImport("kernel32.dll")] private static extern bool GlobalUnlock(IntPtr memory);
    [DllImport("kernel32.dll")] private static extern IntPtr GlobalAlloc(uint flags, UIntPtr size);
    [DllImport("kernel32.dll")] private static extern IntPtr GlobalFree(IntPtr memory);
    [DllImport("kernel32.dll")] private static extern void SetLastError(uint error);

    internal static Dictionary<uint, byte[]> Read()
    {
        var formats = new List<uint>();
        uint current = 0;
        while (true)
        {
            SetLastError(0);
            current = EnumClipboardFormats(current);
            if (current == 0) { if (Marshal.GetLastWin32Error() != 0) throw new InvalidOperationException("enumeration"); break; }
            if (current != 1 && current != 7 && current != 13 && current != 16) throw new NotSupportedException("non-text/private format present; clipboard unchanged");
            formats.Add(current);
            if (formats.Count > 8) throw new InvalidOperationException("format limit");
        }
        var result = new Dictionary<uint, byte[]>();
        foreach (uint format in formats)
        {
            IntPtr handle = GetClipboardData(format);
            ulong size = GlobalSize(handle).ToUInt64();
            if (handle == IntPtr.Zero || size == 0 || size > 65536) throw new InvalidOperationException("data limit");
            IntPtr pointer = GlobalLock(handle);
            if (pointer == IntPtr.Zero) throw new InvalidOperationException("lock");
            try { byte[] data = new byte[(int)size]; Marshal.Copy(pointer, data, 0, data.Length); result.Add(format, data); }
            finally { GlobalUnlock(handle); }
        }
        return result;
    }
    internal static Dictionary<uint, IntPtr> Allocate(Dictionary<uint, byte[]> data)
    {
        var handles = new Dictionary<uint, IntPtr>();
        try
        {
            foreach (var item in data)
            {
                IntPtr handle = GlobalAlloc(2, new UIntPtr((uint)item.Value.Length));
                if (handle == IntPtr.Zero) throw new OutOfMemoryException();
                handles.Add(item.Key, handle);
                IntPtr pointer = GlobalLock(handle);
                if (pointer == IntPtr.Zero) throw new InvalidOperationException("allocation lock");
                try { Marshal.Copy(item.Value, 0, pointer, item.Value.Length); }
                finally { GlobalUnlock(handle); }
            }
            return handles;
        }
        catch { Free(handles); throw; }
    }
    internal static void Free(Dictionary<uint, IntPtr> handles)
    { if (handles != null) foreach (IntPtr value in handles.Values) if (value != IntPtr.Zero) GlobalFree(value); }
    internal static void Write(Dictionary<uint, IntPtr> handles)
    {
        if (!EmptyClipboard()) throw new InvalidOperationException("empty failed");
        foreach (uint format in new List<uint>(handles.Keys))
        {
            if (SetClipboardData(format, handles[format]) == IntPtr.Zero) throw new InvalidOperationException("publish failed");
            handles[format] = IntPtr.Zero;
        }
    }
    [STAThread]
    private static int Main(string[] args)
    {
        int restoreDelay=0;
        if(args.Length==1 && args[0]=="--delay-500")restoreDelay=500;
        else if(args.Length!=0 && !(args.Length==1 && (args[0]=="--synthetic-disposable"||args[0]=="--synthetic-newer-copy")))return 2;
        Dictionary<uint, IntPtr> backup = null, temporary = null;
        bool changed = false;
        uint published = 0;
        using (var owner = new NativeWindowOwner())
        try
        {
            bool simulateNewerCopy = args.Length == 1 && args[0] == "--synthetic-newer-copy";
            if (args.Length == 1 && (args[0] == "--synthetic-disposable" || simulateNewerCopy))
            {
                // Explicit user-authorized replacement of the disposable Notepad copy.
                var fixture = Allocate(new Dictionary<uint, byte[]> { { 13, Encoding.Unicode.GetBytes("clipboard-probe-123\0") } });
                try
                {
                    if (!OpenClipboard(owner.Handle)) { Console.WriteLine("SKIP clipboard busy; fixture not published"); return 2; }
                    try { Write(fixture); }
                    finally { CloseClipboard(); }
                }
                finally { Free(fixture); }
                Console.WriteLine("Published authorized synthetic disposable fixture; original OLE metadata is not restored by this experiment");
            }
            if (!OpenClipboard(owner.Handle)) { Console.WriteLine("SKIP clipboard busy"); return 2; }
            Dictionary<uint, byte[]> snapshot;
            try
            {
                snapshot = Read();
                byte[] unicode;
                if (!snapshot.TryGetValue(13, out unicode) || Encoding.Unicode.GetString(unicode).TrimEnd('\0') != "clipboard-probe-123")
                { Console.WriteLine("SKIP fixture mismatch; no changes"); return 2; }
                backup = Allocate(snapshot);
                temporary = Allocate(new Dictionary<uint, byte[]> { { 13, Encoding.Unicode.GetBytes("clipboard-temporary-validation\0") } });
                changed = true;
                Write(temporary);
            }
            finally { CloseClipboard(); }
            // Windows may synthesize text formats/advance sequence on CloseClipboard.
            published = GetClipboardSequenceNumber();
            if(restoreDelay!=0)System.Threading.Thread.Sleep(restoreDelay);
            if (simulateNewerCopy)
            {
                using (var otherOwner = new NativeWindowOwner())
                {
                    var newer = Allocate(new Dictionary<uint, byte[]> { { 13, Encoding.Unicode.GetBytes("clipboard-probe-newer-copy\0") } });
                    try
                    {
                        if (!OpenClipboard(otherOwner.Handle)) throw new InvalidOperationException("newer-copy lock unavailable");
                        try { Write(newer); }
                        finally { CloseClipboard(); }
                        if (!OpenClipboard(owner.Handle)) throw new InvalidOperationException("race-check lock unavailable");
                        try
                        {
                            if (GetClipboardSequenceNumber() == published || GetClipboardOwner() == owner.Handle) throw new InvalidOperationException("new owner not detected");
                            var retained = Read();
                            if (Encoding.Unicode.GetString(retained[13]).TrimEnd('\0') != "clipboard-probe-newer-copy") throw new InvalidOperationException("new copy altered");
                            Console.WriteLine("PASS simulated different-owner copy detected; old snapshot not restored; newer text retained");
                            changed = false;
                            return 0;
                        }
                        finally { CloseClipboard(); }
                    }
                    finally { Free(newer); }
                }
            }
            if (!OpenClipboard(owner.Handle)) throw new InvalidOperationException("restore lock unavailable");
            try
            {
                if (GetClipboardSequenceNumber() != published || GetClipboardOwner() != owner.Handle) { Console.WriteLine("STOP changed ownership/sequence preserved; no restore overwrite"); changed = false; return 2; }
                Write(backup);
                changed = false;
                var actual = Read();
                foreach (var item in snapshot)
                {
                    byte[] bytes;
                    if (!actual.TryGetValue(item.Key, out bytes) || !Equal(item.Value, bytes)) throw new InvalidOperationException("roundtrip mismatch");
                }
                Console.WriteLine("PASS disposable plain-text clipboard replaced/restored; all snapshot format bytes match");
                return 0;
            }
            finally { CloseClipboard(); }
        }
        catch (NotSupportedException ex) { Console.WriteLine("SKIP " + ex.Message); return 2; }
        catch (Exception ex) { Console.WriteLine("FAIL " + ex.GetType().Name + " restoreMayBeNeeded=" + changed); return 1; }
        finally { Free(backup); Free(temporary); }
    }
    private static bool Equal(byte[] a, byte[] b)
    { if (a.Length != b.Length) return false; for (int i=0;i<a.Length;i++) if (a[i]!=b[i]) return false; return true; }
    private sealed class NativeWindowOwner : NativeWindow, IDisposable
    {
        internal NativeWindowOwner() { CreateHandle(new CreateParams { Caption = "Clipboard validation owner", Parent = new IntPtr(-3) }); }
        public void Dispose() { DestroyHandle(); }
    }
}
