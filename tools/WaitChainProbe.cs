using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

// WCT metadata only: no object names, memory dump, input, COM activation or kill.
internal static class WaitChainProbe
{
    [DllImport("advapi32.dll",SetLastError=true)]private static extern IntPtr OpenThreadWaitChainSession(uint flags,IntPtr callback);
    [DllImport("advapi32.dll")]private static extern void CloseThreadWaitChainSession(IntPtr session);
    [DllImport("advapi32.dll",SetLastError=true)]private static extern bool GetThreadWaitChain(IntPtr session,IntPtr context,uint flags,uint thread,ref uint count,IntPtr nodes,out int cycle);
    private static readonly string[] types={"invalid","critical-section","send-message","mutex","alpc","com","thread-wait","process-wait","thread","com-activation","unknown","socket-io","smb-io"};
    private static int Main(string[] args)
    {
        int pid;if(args.Length!=1||!Int32.TryParse(args[0],out pid))return 2;
        IntPtr session=OpenThreadWaitChainSession(0,IntPtr.Zero);
        if(session==IntPtr.Zero){Console.WriteLine("WCT session unavailable error="+Marshal.GetLastWin32Error());return 1;}
        // WAITCHAIN_NODE_INFO: 8-byte header + aligned 272-byte lock union.
        const int size=280;IntPtr buffer=Marshal.AllocHGlobal(size*16);
        try
        {
            using(var process=Process.GetProcessById(pid))
            {
                if(!process.ProcessName.Equals("ONENOTE",StringComparison.OrdinalIgnoreCase))return 2;
                Console.WriteLine("WCT process="+pid+" threads="+process.Threads.Count);
                foreach(ProcessThread thread in process.Threads)
                {
                    uint count=16;int cycle;
                    if(!GetThreadWaitChain(session,IntPtr.Zero,7,(uint)thread.Id,ref count,buffer,out cycle))
                    {Console.WriteLine("tid="+thread.Id+" unavailable error="+Marshal.GetLastWin32Error());continue;}
                    if(count>16)throw new InvalidOperationException("WCT node limit");
                    string chain="";
                    for(int i=0;i<count;i++)
                    {
                        IntPtr node=IntPtr.Add(buffer,i*size);int type=Marshal.ReadInt32(node),status=Marshal.ReadInt32(node,4);
                        string label=type>=0&&type<types.Length?types[type]:"type-"+type;
                        chain+=(i==0?"":" -> ")+label+"/status-"+status;
                        if(type==8)chain+="(pid="+Marshal.ReadInt32(node,8)+",tid="+Marshal.ReadInt32(node,12)+")";
                    }
                    Console.WriteLine("tid="+thread.Id+" cycle="+(cycle!=0)+" "+chain);
                }
            }
            return 0;
        }
        catch(Exception ex){Console.WriteLine("FAIL "+ex.GetType().Name);return 1;}
        finally{Marshal.FreeHGlobal(buffer);CloseThreadWaitChainSession(session);}
    }
}
