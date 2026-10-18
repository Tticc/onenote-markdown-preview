using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Xml;
using Extensibility;
using Microsoft.Office.Core;
using OneNoteMarkdown.Addin;
internal static class PreviewAudit
{
    private static int Main(string[] args)
    {
        try
        {
            var type=typeof(Addin);
            if(!typeof(IDTExtensibility2).IsAssignableFrom(type)||!typeof(IRibbonExtensibility).IsAssignableFrom(type))throw new Exception("Official interface mismatch");
            if(type.GUID!=new Guid("E036DE73-418B-45F7-973F-A4B09D58E515"))throw new Exception("Unexpected CLSID");
            // Constructing does not call OnConnection, start a thread, or touch clipboard.
            var instance=new Addin();
            var document=new XmlDocument{XmlResolver=null};document.LoadXml(instance.GetCustomUI("audit"));
            var ns=new XmlNamespaceManager(document.NameTable);ns.AddNamespace("r","http://schemas.microsoft.com/office/2009/07/customui");
            foreach(XmlElement control in document.SelectNodes("//*[@onAction]"))
            {var method=type.GetMethod(control.GetAttribute("onAction"));if(method==null||method.GetParameters().Length!=1||method.GetParameters()[0].ParameterType!=typeof(IRibbonControl))throw new Exception("Button signature mismatch");}
            if(type.GetMethod("RibbonLoaded").GetParameters()[0].ParameterType!=typeof(IRibbonUI))throw new Exception("Ribbon onLoad signature mismatch");
            if(args.Length!=0&&Directory.GetFiles(args[0],"*.dll").Length<6)throw new Exception("Dependency package incomplete");
            Console.WriteLine("PASS preview official COM contracts, Ribbon callbacks and dependency layout; not host loading; no registration or input");return 0;
        }
        catch(Exception error){Console.WriteLine("FAIL "+error.GetType().Name+" "+error.Message);return 1;}
    }
}
