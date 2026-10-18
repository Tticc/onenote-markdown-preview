using System;
using System.Runtime.InteropServices;
using System.Text;

// Metadata only. Never requests delayed data or changes ownership/content.
internal static class ClipboardInventory
{
    [DllImport("user32.dll", SetLastError=true)] private static extern bool OpenClipboard(IntPtr window);
    [DllImport("user32.dll")] private static extern bool CloseClipboard();
    [DllImport("user32.dll", SetLastError=true)] private static extern uint EnumClipboardFormats(uint format);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] private static extern int GetClipboardFormatName(uint format, StringBuilder name, int capacity);
    [DllImport("user32.dll")] private static extern uint GetClipboardSequenceNumber();
    [DllImport("user32.dll")] private static extern IntPtr GetClipboardOwner();
    [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint pid);
    [DllImport("kernel32.dll")] private static extern void SetLastError(uint error);
    private static int Main()
    {
        uint before = GetClipboardSequenceNumber();
        if (!OpenClipboard(IntPtr.Zero)) { Console.WriteLine("SKIP clipboard busy; no changes"); return 2; }
        int count = 0;
        try
        {
            uint ownerPid;
            GetWindowThreadProcessId(GetClipboardOwner(), out ownerPid);
            try { using (var process = System.Diagnostics.Process.GetProcessById((int)ownerPid)) Console.WriteLine("Clipboard owner process: " + process.ProcessName); }
            catch (ArgumentException) { Console.WriteLine("Clipboard owner process unavailable"); }
            uint format = 0;
            while (true)
            {
                SetLastError(0);
                format = EnumClipboardFormats(format);
                if (format == 0)
                {
                    if (Marshal.GetLastWin32Error() != 0) throw new InvalidOperationException("Enumeration error");
                    break;
                }
                if (++count > 128) throw new InvalidOperationException("Format count limit");
                string category;
                switch (format)
                {
                    case 1: category = "CF_TEXT"; break;
                    case 2: category = "CF_BITMAP (GDI handle)"; break;
                    case 3: category = "CF_METAFILEPICT (metafile)"; break;
                    case 7: category = "CF_OEMTEXT"; break;
                    case 8: category = "CF_DIB"; break;
                    case 13: category = "CF_UNICODETEXT"; break;
                    case 14: category = "CF_ENHMETAFILE (GDI handle)"; break;
                    case 15: category = "CF_HDROP"; break;
                    case 16: category = "CF_LOCALE"; break;
                    case 17: category = "CF_DIBV5"; break;
                    default:
                        var name = new StringBuilder(256);
                        GetClipboardFormatName(format, name, name.Capacity);
                        string label = name.ToString();
                        category = label == "HTML Format" || label == "Rich Text Format" || label == "PNG" ||
                            label == "CanIncludeInClipboardHistory" || label == "CanUploadToCloudClipboard" || label == "ExcludeClipboardContentFromMonitorProcessing" ||
                            label == "DataObject" || label == "Ole Private Data" || label == "Notepad++ Binary Copy Paste" ||
                            label == "Chromium internal source RFH token" || label == "Chromium Web Custom MIME Data Format" || label == "Chromium internal source URL" ||
                            label == "EnterpriseDataProtectionId" ? label : "other/private (name omitted)";
                        break;
                }
                Console.WriteLine("Clipboard format " + format + ": " + category);
            }
        }
        finally { CloseClipboard(); }
        Console.WriteLine("Format count: " + count);
        Console.WriteLine("Sequence unchanged during inventory: " + (before == GetClipboardSequenceNumber()));
        Console.WriteLine("Inventory only; no data read, snapshot or restore tested");
        return 0;
    }
}
