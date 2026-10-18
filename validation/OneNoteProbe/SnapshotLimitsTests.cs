using System;
internal static class SnapshotLimitsTests
{
    private static int Main()
    {
        var limit=new SnapshotLimits();
        if(limit.Accept(13,0)||limit.Accept(13,-1)||limit.Accept(13,long.MaxValue))return 1;
        if(!limit.Accept(13,4*1024*1024)||limit.Accept(13,1)||!limit.Accept(1,4*1024*1024)||limit.Accept(16,1))return 1;
        if(limit.TotalBytes!=8*1024*1024)return 1;
        limit=new SnapshotLimits();for(uint i=1;i<=128;i++)if(!limit.Accept(i,1))return 1;
        if(limit.Accept(129,1)||limit.TotalBytes!=128)return 1;
        Console.WriteLine("PASS snapshot zero/negative/overflow/per-medium/aggregate/duplicate/format-count boundaries; no clipboard access");return 0;
    }
}
