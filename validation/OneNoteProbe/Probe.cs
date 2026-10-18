using System;
using System.Runtime.InteropServices;
using System.Xml;
using OneNote = Microsoft.Office.Interop.OneNote;

// Read-only capability probe. Never reads page contents or changes the clipboard.
internal static class Probe
{
    [STAThread]
    private static int Main()
    {
        OneNote.Application app = null;
        OneNote.Windows windows = null;
        OneNote.Window window = null;
        try
        {
            app = new OneNote.Application();
            Console.WriteLine("COM activation: OK");
            windows = app.Windows;
            Console.WriteLine("Window count: " + windows.Count);
            if (windows.Count != 1)
            {
                Console.WriteLine("Selection prerequisite: BLOCKED (requires one window)");
                return 2;
            }
            window = windows.CurrentWindow;
            Console.WriteLine("Window handle present: " + (window.WindowHandle != 0));
            Console.WriteLine("Current page present: " + !String.IsNullOrEmpty(window.CurrentPageId));
            Console.WriteLine("Quick note: " + window.SideNote);
            string hierarchy;
            app.GetHierarchy(window.CurrentPageId, OneNote.HierarchyScope.hsSelf, out hierarchy, OneNote.XMLSchema.xs2013);
            var doc = new XmlDocument { XmlResolver = null };
            doc.LoadXml(hierarchy);
            bool fixture = doc.DocumentElement.GetAttribute("name") == "Markdown插件验证";
            Console.WriteLine("Disposable fixture title matches: " + fixture);
            if (!fixture) return 3;
            Console.WriteLine("Window metadata probe: PASS (not a selection or paste test)");
            return 0;
        }
        catch (Exception ex)
        {
            Console.WriteLine("Probe failed: " + ex.GetType().FullName + " HRESULT=0x" + ex.HResult.ToString("X8"));
            return 1;
        }
        finally
        {
            Release(window);
            Release(windows);
            Release(app);
        }
    }

    private static void Release(object value)
    {
        if (value != null && Marshal.IsComObject(value))
            Marshal.ReleaseComObject(value);
    }
}
