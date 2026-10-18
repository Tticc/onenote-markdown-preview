using System;
internal static class ClipboardFormatPolicyTests
{
    private static int Main()
    {
        uint[] ids={13,1,16,8,2,15,50000,50001,50002,50003,50004,50005};
        string[] names={"","","","","","","HTML Format","PNG","Ole Private Data","EnterpriseDataProtectionId","Chromium internal source RFH token","CanUploadToCloudClipboard"};
        string[] expected={"standard-text","standard-text","standard-text","image-dib-needs-validation","image-gdi-needs-validation","reject-file-content","rich-text-needs-validation","image-stream-needs-validation","ole-transport-needs-validation","protection-metadata-needs-validation","reject-unknown","protection-metadata-needs-validation"};
        for(int i=0;i<ids.Length;i++)if(ClipboardFormatPolicy.Classify(ids[i],names[i])!=expected[i])return 1;
        Console.WriteLine("PASS 12 clipboard classification cases; no metadata stripping or clipboard changes");return 0;
    }
}
