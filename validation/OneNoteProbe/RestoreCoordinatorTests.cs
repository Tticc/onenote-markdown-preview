using System;
using System.Threading.Tasks;

internal static class RestoreCoordinatorTests
{
    private sealed class Fake:IRestoreBackend
    {
        internal bool Open,Busy,FailPublish,ChangedOnLock;
        internal int Writes,Closes,Releases;
        internal long CurrentOwner=10;
        internal uint CurrentSequence=2;
        public bool TryOpen(){if(Busy)return false;Open=true;if(ChangedOnLock){CurrentOwner=20;CurrentSequence=3;}return true;}
        public long Owner{get{if(!Open)throw new Exception("owner read outside lock");return CurrentOwner;}}
        public uint Sequence{get{if(!Open)throw new Exception("sequence read outside lock");return CurrentSequence;}}
        public void PublishSnapshot(){if(!Open)throw new Exception("publish outside lock");if(FailPublish)throw new Exception("simulated allocation/publication fault");Writes++;}
        public void Close(){if(!Open)throw new Exception("double close");Open=false;Closes++;}
        public void ReleaseSnapshot(){if(Open)throw new Exception("release while locked");Releases++;}
    }
    private static void Require(bool pass,string name){if(!pass)throw new Exception(name);}
    private static int Main()
    {
        try{Run().GetAwaiter().GetResult();Console.WriteLine("PASS 12 restoration ordering/failure/race cases; no Windows clipboard or input access");return 0;}
        catch(Exception ex){Console.WriteLine("FAIL "+ex.Message);return 1;}
    }
    private static async Task Run()
    {
        var fake=new Fake();int observedDelay=0;
        var gate=new TaskCompletionSource<bool>();
        var coordinator=new RestoreCoordinator(fake,delegate(int delay){Require(!fake.Open,"delay must not hold clipboard");observedDelay=delay;return gate.Task;},10,2);
        Task<string> first=coordinator.AfterSubmission(200);
        Require(!first.IsCompleted&&fake.Writes==0&&observedDelay==200,"asynchronous 200 ms scheduling");
        Require(await coordinator.AfterSubmission(500)=="ignored","no duplicate restore");
        Require(coordinator.CancelBeforeInput()=="ignored","cancel cannot race active delay");
        gate.SetResult(true);
        Require(await first=="restored"&&fake.Writes==1&&fake.Closes==1&&fake.Releases==1&&!coordinator.SnapshotRetained,"successful cleanup");
        Require(coordinator.RetryCleanupOnly()=="ignored"&&fake.Writes==1,"terminal operation not repeated");

        fake=new Fake();coordinator=new RestoreCoordinator(fake,delegate(int delay){observedDelay=delay;return Task.FromResult(true);},10,2);
        Require(await coordinator.AfterSubmission(500)=="restored"&&observedDelay==500,"500 ms selected");
        fake=new Fake();coordinator=new RestoreCoordinator(fake,delegate(int delay){observedDelay=delay;return Task.FromResult(true);},10,2);
        Require(await coordinator.AfterSubmission(-1)=="restored"&&observedDelay==200,"invalid config fallback");

        fake=new Fake{Busy=true};coordinator=new RestoreCoordinator(fake,delay=>Task.FromResult(true),10,2);
        Require(await coordinator.AfterSubmission(200)=="busy-snapshot-retained"&&coordinator.SnapshotRetained&&fake.Writes==0&&fake.Releases==0,"busy retains original");
        fake.Busy=false;Require(coordinator.RetryCleanupOnly()=="restored"&&fake.Writes==1,"explicit cleanup recovery");

        fake=new Fake{ChangedOnLock=true};coordinator=new RestoreCoordinator(fake,delay=>Task.FromResult(true),10,2);
        Require(await coordinator.AfterSubmission(200)=="newer-content-preserved"&&fake.Writes==0&&fake.Releases==1,"locked race protection");

        fake=new Fake{FailPublish=true};coordinator=new RestoreCoordinator(fake,delay=>Task.FromResult(true),10,2);
        Require(await coordinator.AfterSubmission(200)=="publish-failed-snapshot-retained"&&coordinator.SnapshotRetained&&!fake.Open&&fake.Releases==0,"failed publish retains original");
        fake.CurrentSequence=3;fake.FailPublish=false;
        Require(coordinator.RetryCleanupOnly()=="newer-content-preserved"&&fake.Writes==0,"failure recovery cannot overwrite later copy");
    }
}
