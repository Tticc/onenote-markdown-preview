using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Xml;
using System.Text;
using OneNote = Microsoft.Office.Interop.OneNote;

// Read-only fixture probe. Outputs structure and numeric boundaries, never note text/IDs.
internal static class SelectionProbe
{
    [StructLayout(LayoutKind.Sequential)]
    private struct Rect { public int Left, Top, Right, Bottom; }
    [StructLayout(LayoutKind.Sequential)]
    private struct GuiInfo
    {
        public int Size, Flags;
        public IntPtr Active, Focus, Capture, MenuOwner, MoveSize, Caret;
        public Rect CaretRect;
    }
    [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint pid);
    [DllImport("user32.dll")] private static extern bool GetGUIThreadInfo(uint thread, ref GuiInfo info);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr window, StringBuilder name, int count);
    [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] private static extern IntPtr GetAncestor(IntPtr window, uint flag);
    [STAThread]
    private static int Main()
    {
        bool visible = false;
        foreach (var process in Process.GetProcessesByName("ONENOTE"))
        {
            using (process) { if (process.MainWindowHandle != IntPtr.Zero) visible = true; }
        }
        if (!visible) { Console.WriteLine("SKIP: open the disposable fixture in OneNote first; no COM activation attempted"); return 2; }
        OneNote.Application app = null;
        OneNote.Windows windows = null;
        OneNote.Window window = null;
        try
        {
            app = new OneNote.Application();
            windows = app.Windows;
            if (windows.Count != 1) { Console.WriteLine("SKIP: requires one window"); return 2; }
            window = windows.CurrentWindow;
            Console.WriteLine("Full page view: " + window.FullPageView);
            Console.WriteLine("Quick note view: " + window.SideNote);
            uint pid;
            uint thread = GetWindowThreadProcessId(new IntPtr((long)window.WindowHandle), out pid);
            var gui = new GuiInfo { Size = Marshal.SizeOf(typeof(GuiInfo)) };
            string focusClass = "";
            if (GetGUIThreadInfo(thread, ref gui))
            {
                var name = new StringBuilder(256);
                GetClassName(gui.Focus, name, name.Capacity);
                focusClass = name.ToString();
                Console.WriteLine("Host GUI thread focus class: " + name.ToString());
                Console.WriteLine("Host GUI thread has caret: " + (gui.Caret != IntPtr.Zero));
            }
            else Console.WriteLine("Host GUI thread focus unavailable");
            string page = window.CurrentPageId;
            string hierarchy;
            app.GetHierarchy(page, OneNote.HierarchyScope.hsSelf, out hierarchy, OneNote.XMLSchema.xs2013);
            var metadata = new XmlDocument { XmlResolver = null };
            metadata.LoadXml(hierarchy);
            if (metadata.DocumentElement.GetAttribute("name") != "Markdown插件验证")
            { Console.WriteLine("SKIP: disposable fixture title mismatch"); return 2; }
            string content;
            var elapsed = Stopwatch.StartNew();
            app.GetPageContent(page, out content, OneNote.PageInfo.piSelection, OneNote.XMLSchema.xs2013);
            Console.WriteLine("Selection export milliseconds: " + elapsed.ElapsedMilliseconds);
            if (content.Length > 1000000) { Console.WriteLine("SKIP: fixture exceeds probe limit"); return 2; }
            var xml = SelectionScreen.Parse(content);
            Console.WriteLine("Reusable structural screening: " + SelectionScreen.Evaluate(xml));
            long host = (long)window.WindowHandle;
            long hostRoot = GetAncestor(new IntPtr(host), 2).ToInt64();
            IntPtr foreground = GetForegroundWindow();
            IntPtr foregroundRoot = GetAncestor(foreground, 2);
            uint foregroundPid;
            GetWindowThreadProcessId(foregroundRoot, out foregroundPid);
            Console.WriteLine("Foreground process is OneNote host: " + (foregroundPid == pid));
            Console.WriteLine("Foreground root equals COM window: " + (foregroundRoot.ToInt64() == host));
            Console.WriteLine("Foreground root equals COM window root: " + (foregroundRoot.ToInt64() == hostRoot));
            Console.WriteLine("Host focus handle present: " + (gui.Focus != IntPtr.Zero));
            var initialEvidence = SelectionEvidence.Capture(content, page, hostRoot, gui.Focus.ToInt64(), GetAncestor(GetForegroundWindow(), 2).ToInt64(), focusClass, (int)windows.Count);
            if (initialEvidence == null) Console.WriteLine("Evidence revalidation: refused initial environment or structure");
            else
            {
                string secondContent;
                app.GetPageContent(window.CurrentPageId, out secondContent, OneNote.PageInfo.piSelection, OneNote.XMLSchema.xs2013);
                var secondGui = new GuiInfo { Size = Marshal.SizeOf(typeof(GuiInfo)) };
                string secondClass = "";
                if (GetGUIThreadInfo(thread, ref secondGui))
                {
                    var name = new StringBuilder(256);
                    GetClassName(secondGui.Focus, name, name.Capacity);
                    secondClass = name.ToString();
                }
                var secondEvidence = SelectionEvidence.Capture(secondContent, window.CurrentPageId, GetAncestor(new IntPtr((long)window.WindowHandle), 2).ToInt64(),
                    secondGui.Focus.ToInt64(), GetAncestor(GetForegroundWindow(), 2).ToInt64(), secondClass, (int)windows.Count);
                Console.WriteLine("Conservative double-read evidence matches: " + initialEvidence.Matches(secondEvidence));
            }
            var ns = new XmlNamespaceManager(xml.NameTable);
            ns.AddNamespace("one", "http://schemas.microsoft.com/office/onenote/2013/onenote");
            int selectedOutlines = xml.SelectNodes("//one:Outline[@selected='all' or @selected='partial']", ns).Count;
            int selectedTitles = xml.SelectNodes("//one:Title[@selected='all' or @selected='partial']", ns).Count;
            int selectedCells = xml.SelectNodes("//one:Cell[@selected='all' or @selected='partial']", ns).Count;
            int selectedObjects = xml.SelectNodes("//*[@selected='all' or @selected='partial'][self::one:Image or self::one:InkDrawing or self::one:InsertedFile or self::one:Table]", ns).Count;
            Console.WriteLine("Selected Outline count: " + selectedOutlines);
            Console.WriteLine("Selected Cell count: " + selectedCells);
            int decodedLength = 0;
            bool textDecodable = true;
            foreach (XmlElement textElement in xml.SelectNodes("//one:T[@selected='all' or @selected='partial']", ns))
            {
                try
                {
                    var fragment = new XmlDocument { XmlResolver = null };
                    fragment.LoadXml("<root>" + textElement.InnerText.Replace("&nbsp;", "&#160;") + "</root>");
                    decodedLength += fragment.DocumentElement.InnerText.Length;
                }
                catch (XmlException) { textDecodable = false; }
            }
            if (selectedTitles != 0) Console.WriteLine("Structural refusal: title selection");
            else if (selectedOutlines != 1) Console.WriteLine("Structural refusal: requires exactly one selected body container");
            else if (selectedCells > 1) Console.WriteLine("Structural refusal: cross-cell selection");
            else if (selectedObjects != 0) Console.WriteLine("Structural refusal: selected non-text object");
            else if (!textDecodable) Console.WriteLine("Structural refusal: text needs unsupported diagnostic decoding");
            else if (decodedLength == 0) Console.WriteLine("Structural refusal: no selected text characters (paragraph separators not classified)");
            else Console.WriteLine("Structural screening only: further focus/range/content checks required");
            foreach (string name in new[] { "Outline", "OE", "T", "Image", "InkDrawing", "InsertedFile", "Table" })
                Console.WriteLine("Element " + name + " count: " + xml.SelectNodes("//one:" + name, ns).Count);
            int selected = 0;
            foreach (XmlElement element in xml.SelectNodes("//*"))
            {
                foreach (XmlAttribute attribute in element.Attributes)
                {
                    if (attribute.LocalName.IndexOf("select", StringComparison.OrdinalIgnoreCase) < 0) continue;
                    selected++;
                    // Attribute contents are not trusted to be numeric; never print arbitrary strings.
                    int value;
                    bool flag;
                    string safe = Int32.TryParse(attribute.Value, out value) ? value.ToString() :
                        (Boolean.TryParse(attribute.Value, out flag) ? flag.ToString() : "[non-numeric value omitted]");
                    if (attribute.Value == "all" || attribute.Value == "partial" || attribute.Value == "none") safe = attribute.Value;
                    Console.WriteLine("Selection marker element=" + element.LocalName + " attribute=" + attribute.LocalName + " value=" + safe);
                    if (element.LocalName == "T")
                    {
                        Console.WriteLine("Selected T character count: " + element.InnerText.Length);
                        Console.WriteLine("Selected T exactly matches controlled fixture token: " + (element.InnerText == "**测试文字**"));
                        XmlNode parent = element.ParentNode;
                        Console.WriteLine("Sibling T count in same OE: " + parent.SelectNodes("one:T", ns).Count);
                        int index = 0;
                        int prefix = 0;
                        foreach (XmlNode sibling in parent.ChildNodes)
                        {
                            if (Object.ReferenceEquals(sibling, element)) break;
                            if (sibling.LocalName == "T") { index++; prefix += sibling.InnerText.Length; }
                        }
                        Console.WriteLine("Selected T index within OE: " + index);
                        Console.WriteLine("Preceding T serialized character count: " + prefix);
                        Console.WriteLine("OE has object identity: " + (parent.Attributes != null && parent.Attributes["objectID"] != null));
                        Console.WriteLine("Selected T contains serialized markup: " + element.InnerText.Contains("<"));
                        try
                        {
                            var fragment = new XmlDocument { XmlResolver = null };
                            // Diagnostic only: not an HTML sanitizer or a production renderer.
                            fragment.LoadXml("<root>" + element.InnerText.Replace("&nbsp;", "&#160;") + "</root>");
                            string text = fragment.DocumentElement.InnerText;
                            Console.WriteLine("XML-compatible text fragment parsed: True");
                            Console.WriteLine("Decoded selected text character count: " + text.Length);
                            Console.WriteLine("Decoded selection matches controlled token: " + (text == "**测试文字**"));
                            Console.WriteLine("Fragment child element count: " + fragment.SelectNodes("/root//*").Count);
                        }
                        catch (XmlException ex)
                        {
                            Console.WriteLine("XML-compatible text fragment parsed: False; error position=" + ex.LinePosition);
                            string markup = element.InnerText;
                            Console.WriteLine("Contains HTML-only font tag: " + markup.Contains("<font"));
                            Console.WriteLine("Contains HTML br without XML slash: " + markup.Contains("<br>"));
                            Console.WriteLine("Contains HTML-only entity: " + System.Text.RegularExpressions.Regex.IsMatch(markup, @"&(?!amp;|lt;|gt;|quot;|apos;|nbsp;|#)[A-Za-z]+;"));
                            Console.WriteLine("Contains unquoted attribute: " + System.Text.RegularExpressions.Regex.IsMatch(markup, @"\s[A-Za-z-]+=[^\s'""<>][^\s<>]*"));
                        }
                    }
                }
            }
            Console.WriteLine("Selection marker count: " + selected);
            Console.WriteLine("Read-only fixture export complete; precision not yet established");
            return 0;
        }
        catch (Exception ex)
        {
            Console.WriteLine("FAIL " + ex.GetType().Name + " HRESULT=0x" + ex.HResult.ToString("X8"));
            return 1;
        }
        finally
        {
            foreach (object obj in new object[] { window, windows, app })
                if (obj != null && Marshal.IsComObject(obj)) Marshal.ReleaseComObject(obj);
        }
    }
}
