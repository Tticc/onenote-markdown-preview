// Classification experiment only. Auxiliary transport approval is deliberately
// separate from classification; no data is dropped by this class.
internal static class ClipboardFormatPolicy
{
    internal static string Classify(uint id,string name)
    {
        if(id==1||id==7||id==13||id==16)return "standard-text";
        if(id==8||id==17)return "image-dib-needs-validation";
        if(id==2||id==3||id==14)return "image-gdi-needs-validation";
        if(id==15)return "reject-file-content";
        if(name=="HTML Format"||name=="Rich Text Format")return "rich-text-needs-validation";
        if(name=="PNG")return "image-stream-needs-validation";
        if(name=="DataObject"||name=="Ole Private Data")return "ole-transport-needs-validation";
        if(name=="EnterpriseDataProtectionId"||name=="CanIncludeInClipboardHistory"||name=="CanUploadToCloudClipboard"||name=="ExcludeClipboardContentFromMonitorProcessing")return "protection-metadata-needs-validation";
        return "reject-unknown";
    }
}
