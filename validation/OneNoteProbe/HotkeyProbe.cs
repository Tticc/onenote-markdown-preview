using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

// 90-second observational keyboard fixture: NEVER suppresses or injects keys.
internal static class HotkeyProbe
{
    private delegate IntPtr Hook(int code,IntPtr message,IntPtr data);
    [DllImport("user32.dll",SetLastError=true)]private static extern IntPtr SetWindowsHookEx(int id,Hook hook,IntPtr module,uint thread);
    [DllImport("user32.dll")]private static extern bool UnhookWindowsHookEx(IntPtr hook);
    [DllImport("user32.dll")]private static extern IntPtr CallNextHookEx(IntPtr hook,int code,IntPtr message,IntPtr data);
    [DllImport("user32.dll")]private static extern short GetAsyncKeyState(int key);
    [DllImport("user32.dll")]private static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")]private static extern uint GetWindowThreadProcessId(IntPtr window,out uint pid);
    [DllImport("user32.dll")]private static extern bool GetGUIThreadInfo(uint thread,ref Gui info);
    [DllImport("user32.dll",CharSet=CharSet.Unicode)]private static extern int GetClassName(IntPtr window,StringBuilder name,int count);
    [DllImport("kernel32.dll",CharSet=CharSet.Unicode)]private static extern IntPtr GetModuleHandle(string module);
    [StructLayout(LayoutKind.Sequential)]private struct Rect{public int A,B,C,D;}
    [StructLayout(LayoutKind.Sequential)]private struct Gui{public int Size,Flags;public IntPtr Active,Focus,Capture,Menu,Move,Caret;public Rect R;}
    private static readonly Hook callback=OnKey;
    private static IntPtr handle;
    private static Control dispatcher;
    private static string log;
    private static bool down;
    private static IntPtr OnKey(int code,IntPtr message,IntPtr data)
    {
        try
        {
            if(code>=0&&Marshal.ReadInt32(data)==0x44)
            {
                int msg=message.ToInt32();
                if(msg==0x101||msg==0x105)down=false;
                if((msg==0x100||msg==0x104)&&!down)
                {
                    down=true;
                    if(GetAsyncKeyState(0x11)<0&&GetAsyncKeyState(0x10)<0&&(Marshal.ReadInt32(data,8)&0x10)==0)
                    {
                        uint pid;uint thread=GetWindowThreadProcessId(GetForegroundWindow(),out pid);
                        var gui=new Gui{Size=Marshal.SizeOf(typeof(Gui))};
                        var name=new StringBuilder(256);
                        if(GetGUIThreadInfo(thread,ref gui))GetClassName(gui.Focus,name,256);
                        uint capturedPid=pid;string capturedClass=name.ToString();
                        dispatcher.BeginInvoke((Action)delegate
                        {
                            string processName="unavailable";
                            try{using(var p=Process.GetProcessById((int)capturedPid))processName=p.ProcessName;}catch(ArgumentException){}
                            string classification=processName.Equals("ONENOTE",StringComparison.OrdinalIgnoreCase)&&capturedClass=="OneNote::DocumentCanvas"?"body-focus-candidate":"non-target";
                            File.AppendAllText(log,DateTime.UtcNow.ToString("o")+" CtrlShiftD "+classification+" passed-through; no selection claim"+Environment.NewLine);
                        });
                    }
                }
            }
        }
        catch { /* Observational callback always passes through. */ }
        return CallNextHookEx(handle,code,message,data);
    }
    [STAThread]private static int Main(string[] args)
    {
        if(args.Length!=1)return 2;log=args[0];
        using(dispatcher=new Control())
        using(var timer=new System.Windows.Forms.Timer{Interval=90000})
        {
            dispatcher.CreateControl();
            handle=SetWindowsHookEx(13,callback,GetModuleHandle(null),0);
            if(handle==IntPtr.Zero)return 1;
            try
            {
                File.AppendAllText(log,DateTime.UtcNow.ToString("o")+" START observational hook; all keys pass through; 90 second lifetime"+Environment.NewLine);
                timer.Tick+=delegate{Application.ExitThread();};timer.Start();Application.Run();
                File.AppendAllText(log,DateTime.UtcNow.ToString("o")+" STOP"+Environment.NewLine);return 0;
            }
            finally{UnhookWindowsHookEx(handle);handle=IntPtr.Zero;}
        }
    }
}
