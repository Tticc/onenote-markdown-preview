using System;
using System.Security.Cryptography;
using System.Text;

// Conservative experiment: entire export equality detects range/content changes but
// may also cancel on harmless metadata changes. Never grants permission to paste.
internal sealed class SelectionEvidence
{
    private readonly byte[] digest;
    private readonly string page;
    private readonly long window, focus;
    private SelectionEvidence(string pageId, long hwnd, long focusHwnd, byte[] hash)
    { page = pageId; window = hwnd; focus = focusHwnd; digest = hash; }

    internal static SelectionEvidence Capture(string xml, string pageId, long hwnd, long focusHwnd,
        long foregroundRoot, string focusClass, int windowCount)
    {
        if (windowCount != 1 || hwnd == 0 || focusHwnd == 0 || foregroundRoot != hwnd ||
            focusClass != "OneNote::DocumentCanvas" || String.IsNullOrEmpty(pageId)) return null;
        if (SelectionScreen.Evaluate(SelectionScreen.Parse(xml)) != "candidate-needs-focus-range-style-checks") return null;
        using (var sha = SHA256.Create())
            return new SelectionEvidence(pageId, hwnd, focusHwnd, sha.ComputeHash(Encoding.UTF8.GetBytes(xml)));
    }

    internal bool Matches(SelectionEvidence other)
    {
        if (other == null || page != other.page || window != other.window || focus != other.focus) return false;
        int difference = 0;
        for (int i = 0; i < digest.Length; i++) difference |= digest[i] ^ other.digest[i];
        return difference == 0;
    }
}
