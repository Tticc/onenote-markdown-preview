internal static class RestorePolicy
{
    internal static int Delay(int configured) { return configured == 500 ? 500 : 200; }
    // Must be called while the clipboard is locked by the restoring owner.
    internal static bool MayRestore(long expectedOwner,long currentOwner,uint expectedSequence,uint currentSequence)
    { return expectedOwner!=0 && expectedOwner==currentOwner && expectedSequence==currentSequence; }
}
