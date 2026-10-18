using System;
using System.Runtime.InteropServices;
using System.Xml;
using Extensibility;
using Microsoft.Office.Core;

// Managed/offline tests only. Does not activate COM or contact OneNote.
internal static class ContractTests
{
    private static void Check(bool value, string message)
    {
        if (!value) throw new InvalidOperationException(message);
        Console.WriteLine("PASS " + message);
    }

    private static int Main()
    {
        try
        {
            Type implementation = typeof(ProbeAddin);
            Check(typeof(IDTExtensibility2).IsAssignableFrom(implementation), "Official lifecycle interface identity");
            Check(typeof(IRibbonExtensibility).IsAssignableFrom(implementation), "Official Ribbon interface identity");
            foreach (Type contract in new[] { typeof(IDTExtensibility2), typeof(IRibbonExtensibility) })
            {
                Check(contract.IsImport, "Imported contract: " + contract.Name);
                var flags = (TypeLibTypeAttribute)Attribute.GetCustomAttribute(contract, typeof(TypeLibTypeAttribute));
                Check(flags != null && (flags.Value & TypeLibTypeFlags.FDual) != 0, "Dual contract: " + contract.Name);
                Check(implementation.GetInterfaceMap(contract).TargetMethods.Length == contract.GetMethods().Length, "Complete method map: " + contract.Name);
            }
            Check(implementation.GetMethod("VerifyCallback").GetParameters()[0].ParameterType == typeof(IRibbonControl), "Official Ribbon callback parameter");
            var instance = new ProbeAddin();
            Array custom = new object[0];
            IDTExtensibility2 lifecycle = instance;
            lifecycle.OnConnection(null, ext_ConnectMode.ext_cm_Startup, null, ref custom);
            lifecycle.OnAddInsUpdate(ref custom);
            lifecycle.OnStartupComplete(ref custom);
            var xml = new XmlDocument { XmlResolver = null };
            xml.LoadXml(((IRibbonExtensibility)instance).GetCustomUI("offline"));
            var ns = new XmlNamespaceManager(xml.NameTable);
            ns.AddNamespace("r", "http://schemas.microsoft.com/office/2009/07/customui");
            Check(xml.SelectNodes("//r:button[@onAction='VerifyCallback']", ns).Count == 1, "Well-formed Ribbon with callback (not Office schema validation)");
            instance.VerifyCallback(null);
            lifecycle.OnBeginShutdown(ref custom);
            lifecycle.OnDisconnection(ext_DisconnectMode.ext_dm_HostShutdown, ref custom);
            Console.WriteLine("Offline lifecycle sequence completed. Not a host/marshaling test.");
            return 0;
        }
        catch (Exception ex)
        {
            Console.WriteLine("FAIL " + ex.GetType().Name + " " + ex.Message);
            return 1;
        }
    }
}
