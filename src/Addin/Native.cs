using System;
using System.Runtime.InteropServices;
using System.Text;
namespace OneNoteMarkdown.Addin
{
    internal static class Native
    {
        [StructLayout(LayoutKind.Sequential)]internal struct Rect{public int A,B,C,D;}
        [StructLayout(LayoutKind.Sequential)]internal struct Gui{public int Size,Flags;public IntPtr Active,Focus,Capture,Menu,Move,Caret;public Rect Rect;}
        [StructLayout(LayoutKind.Sequential)]internal struct Input{public uint Type;public Union Data;}
        [StructLayout(LayoutKind.Explicit)]internal struct Union{[FieldOffset(0)]public Key Keyboard;[FieldOffset(0)]public Mouse Mouse;}
        [StructLayout(LayoutKind.Sequential)]internal struct Key{public ushort Code,Scan;public uint Flags,Time;public UIntPtr Extra;}
        [StructLayout(LayoutKind.Sequential)]internal struct Mouse{public int X,Y;public uint Data,Flags,Time;public UIntPtr Extra;}
        internal delegate IntPtr Hook(int code,IntPtr message,IntPtr data);
        [DllImport("user32.dll")]internal static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")]internal static extern IntPtr GetAncestor(IntPtr w,uint flag);
        [DllImport("user32.dll")]internal static extern uint GetWindowThreadProcessId(IntPtr w,out uint pid);
        [DllImport("user32.dll")]internal static extern bool GetGUIThreadInfo(uint thread,ref Gui info);
        [DllImport("user32.dll",CharSet=CharSet.Unicode)]internal static extern int GetClassName(IntPtr w,StringBuilder b,int length);
        [DllImport("user32.dll")]internal static extern short GetAsyncKeyState(int code);
        [DllImport("user32.dll")]internal static extern uint SendInput(uint count,Input[] values,int size);
        [DllImport("user32.dll")]internal static extern IntPtr SetWindowsHookEx(int id,Hook callback,IntPtr module,uint thread);
        [DllImport("user32.dll")]internal static extern bool UnhookWindowsHookEx(IntPtr hook);
        [DllImport("user32.dll")]internal static extern IntPtr CallNextHookEx(IntPtr hook,int code,IntPtr message,IntPtr data);
        [DllImport("kernel32.dll",CharSet=CharSet.Unicode)]internal static extern IntPtr GetModuleHandle(string name);
        internal static bool BodyFocus(out IntPtr root,out IntPtr focus)
        {
            root=GetAncestor(GetForegroundWindow(),2);focus=IntPtr.Zero;uint pid;
            uint thread=GetWindowThreadProcessId(root,out pid);
            var gui=new Gui{Size=Marshal.SizeOf(typeof(Gui))};
            if(!GetGUIThreadInfo(thread,ref gui))return false;
            focus=gui.Focus;var name=new StringBuilder(256);GetClassName(focus,name,256);
            if(name.ToString()!="OneNote::DocumentCanvas")return false;
            try{using(var p=System.Diagnostics.Process.GetProcessById((int)pid))return p.ProcessName.Equals("ONENOTE",StringComparison.OrdinalIgnoreCase);}catch(ArgumentException){return false;}
        }
        internal static uint Paste()
        {
            var inputs=new Input[4];for(int i=0;i<4;i++)inputs[i].Type=1;
            inputs[0].Data.Keyboard.Code=0x11;inputs[1].Data.Keyboard.Code=0x56;
            inputs[2].Data.Keyboard.Code=0x56;inputs[2].Data.Keyboard.Flags=2;
            inputs[3].Data.Keyboard.Code=0x11;inputs[3].Data.Keyboard.Flags=2;
            return SendInput(4,inputs,Marshal.SizeOf(typeof(Input)));
        }
    }
}
