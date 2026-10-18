using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using Extensibility;
using Microsoft.Office.Core;

// Temporary loading probe only: no input injection, clipboard access or document writes.
[ComVisible(true), Guid("AE4143BE-ABDB-4B51-989B-D894A7BD63D5"), ProgId("OneNoteMarkdown.ValidationProbe"), ClassInterface(ClassInterfaceType.AutoDispatch)]
public sealed class ProbeAddin : IDTExtensibility2, IRibbonExtensibility
{
    private static readonly Stopwatch Lifetime = Stopwatch.StartNew();
    public void OnConnection(object application, ext_ConnectMode connectMode, object addin, ref Array custom) { Log("Connected mode=" + connectMode + " applicationPresent=" + (application != null) + " addinPresent=" + (addin != null)); }
    public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom) { Log("Disconnected mode=" + removeMode); }
    public void OnAddInsUpdate(ref Array custom) { }
    public void OnStartupComplete(ref Array custom) { Log("StartupComplete"); }
    public void OnBeginShutdown(ref Array custom) { Log("BeginShutdown"); }
    public string GetCustomUI(string ribbonId)
    {
        Log("GetCustomUI");
        return "<customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\"><ribbon><tabs><tab id=\"MarkdownValidationTab\" label=\"Markdown Probe\"><group id=\"MarkdownValidationGroup\" label=\"Validation only\"><button id=\"MarkdownValidationPing\" label=\"Verify callback\" onAction=\"VerifyCallback\" screentip=\"Records a callback without changing notes\"/></group></tab></tabs></ribbon></customUI>";
    }
    public void VerifyCallback(IRibbonControl control)
    {
        try { Log("RibbonCallback PASS controlPresent=" + (control != null) + " controlIdPresent=" + (control != null && !String.IsNullOrEmpty(control.Id))); }
        catch (Exception ex) { Log("RibbonCallback failure HRESULT=0x" + ex.HResult.ToString("X8")); }
    }
    private static void Log(string message)
    {
        try
        {
            string folder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OneNoteMarkdownValidation");
            Directory.CreateDirectory(folder);
            File.AppendAllText(Path.Combine(folder, "probe.log"), DateTime.UtcNow.ToString("o") + " pid=" + Process.GetCurrentProcess().Id + " elapsedMs=" + Lifetime.ElapsedMilliseconds + " " + message + Environment.NewLine);
        }
        catch { /* A diagnostic failure must not escape a host callback. */ }
    }
}
