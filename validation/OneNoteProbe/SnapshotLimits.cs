using System;
using System.Collections.Generic;

internal sealed class SnapshotLimits
{
    private readonly HashSet<uint> formats=new HashSet<uint>();
    internal long TotalBytes { get; private set; }
    internal bool Accept(uint id,long bytes)
    {
        if(formats.Contains(id)||formats.Count>=128||bytes<=0||bytes>4*1024*1024||TotalBytes>8*1024*1024-bytes)return false;
        formats.Add(id);TotalBytes+=bytes;return true;
    }
}
