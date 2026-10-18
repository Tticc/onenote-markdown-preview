using System;
internal static class RestorePolicyTests
{
    private static int Main()
    {
        int[] inputs={200,500,0,-1,201,1000};int[] expected={200,500,200,200,200,200};
        for(int i=0;i<inputs.Length;i++)if(RestorePolicy.Delay(inputs[i])!=expected[i])return 1;
        if(!RestorePolicy.MayRestore(10,10,3,3))return 1;
        if(RestorePolicy.MayRestore(10,11,3,3)||RestorePolicy.MayRestore(10,10,3,4)||RestorePolicy.MayRestore(0,0,3,3))return 1;
        Console.WriteLine("PASS 10 delay/default/ownership policy cases; no input or clipboard changes");return 0;
    }
}
