using System;
namespace OneNoteMarkdown
{
    public sealed class OperationState
    {
        private readonly object gate=new object();private long generation;private bool enabled,busy,faulted;
        public bool Enabled{get{lock(gate)return enabled&&!faulted;}}
        public bool Busy{get{lock(gate)return busy;}}
        public bool Enable(){lock(gate){if(busy)return false;enabled=true;faulted=false;return true;}}
        public void Disable(){lock(gate){enabled=false;generation++;}}
        public long Begin(){lock(gate){if(!enabled||faulted||busy)return 0;busy=true;return ++generation;}}
        public bool Valid(long ticket){lock(gate)return ticket!=0&&ticket==generation&&enabled&&!faulted&&busy;}
        public void Fault(){lock(gate){enabled=false;faulted=true;generation++;}}
        public void End(){lock(gate)busy=false;}
        public static int Delay(int value){return value==500?500:200;}
    }
}
