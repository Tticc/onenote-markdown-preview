using System;
using System.Threading.Tasks;

// Testable restoration ordering. Backend owns a PREALLOCATED independent snapshot.
// No Windows implementation/production support is implied by passing these tests.
internal interface IRestoreBackend
{
    bool TryOpen();
    long Owner { get; }
    uint Sequence { get; }
    void PublishSnapshot();
    void Close();
    void ReleaseSnapshot();
}

internal sealed class RestoreCoordinator
{
    private readonly IRestoreBackend backend;
    private readonly Func<int,Task> wait;
    private readonly long owner;
    private readonly uint sequence;
    private bool inFlight, released, terminal;
    internal RestoreCoordinator(IRestoreBackend backend,Func<int,Task> wait,long owner,uint sequence)
    { this.backend=backend;this.wait=wait;this.owner=owner;this.sequence=sequence; }
    internal bool SnapshotRetained { get { return !released; } }
    private void Release()
    {
        if(!released){backend.ReleaseSnapshot();released=true;}
        terminal=true;
    }
    internal async Task<string> AfterSubmission(int delay)
    {
        if(inFlight||terminal)return "ignored";
        inFlight=true;
        try
        {
            // Explicitly no clipboard lock during empirical delay.
            await wait(RestorePolicy.Delay(delay));
            return AttemptLocked();
        }
        finally{inFlight=false;}
    }
    internal string CancelBeforeInput()
    {
        if(inFlight||terminal)return "ignored";
        inFlight=true;
        try{return AttemptLocked();}finally{inFlight=false;}
    }
    internal string RetryCleanupOnly()
    {
        if(inFlight||terminal)return "ignored";
        inFlight=true;
        try{return AttemptLocked();}finally{inFlight=false;}
    }
    private string AttemptLocked()
    {
        if(!backend.TryOpen())return "busy-snapshot-retained";
        string outcome;
        try
        {
            if(!RestorePolicy.MayRestore(owner,backend.Owner,sequence,backend.Sequence))outcome="newer-content-preserved";
            else
            {
                try{backend.PublishSnapshot();outcome="restored";}
                catch(Exception){return "publish-failed-snapshot-retained";}
            }
        }
        finally{backend.Close();}
        Release();
        return outcome;
    }
}
