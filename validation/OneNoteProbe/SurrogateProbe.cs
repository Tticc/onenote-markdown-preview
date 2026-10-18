using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using Extensibility;

internal static class SurrogateProbe
{
    [DllImport("ole32.dll", PreserveSig = true)]
    private static extern int CoCreateInstance(ref Guid clsid, IntPtr outer, uint context, ref Guid iid, out IntPtr instance);

    [STAThread]
    private static int Main()
    {
        Guid clsid = new Guid("AE4143BE-ABDB-4B51-989B-D894A7BD63D5");
        Guid iid = new Guid("000C0396-0000-0000-C000-000000000046");
        IntPtr instance;
        int hr = CoCreateInstance(ref clsid, IntPtr.Zero, 4, ref iid, out instance);
        Console.WriteLine("CLSCTX_LOCAL_SERVER Ribbon interface: 0x" + hr.ToString("X8"));
        if (hr < 0) return 1;
        try
        {
            object obj = Marshal.GetObjectForIUnknown(instance);
            try
            {
                string xml = ((IRibbonExtensibility)obj).GetCustomUI("probe");
                if (!xml.Contains("Markdown Probe")) return 2;
                Console.WriteLine("Cross-process GetCustomUI: True");
                var lifecycle = (IDTExtensibility2)obj;
                Array custom = new object[] { "activation-fixture" };
                lifecycle.OnConnection(null, ext_ConnectMode.ext_cm_External, null, ref custom);
                lifecycle.OnStartupComplete(ref custom);
                lifecycle.OnBeginShutdown(ref custom);
                lifecycle.OnDisconnection(ext_DisconnectMode.ext_dm_UserClosed, ref custom);
                Console.WriteLine("Official lifecycle activation/release sequence: PASS");
            }
            finally { if (Marshal.IsComObject(obj)) Marshal.ReleaseComObject(obj); }
        }
        finally { Marshal.Release(instance); }
        return 0;
    }
}
