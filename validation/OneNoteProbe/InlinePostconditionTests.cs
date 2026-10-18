using System;

internal static class InlinePostconditionTests
{
    private static string Page(string body) { return "<Page xmlns='http://schemas.microsoft.com/office/onenote/2013/onenote'><Outline><OEChildren>" + body + "</OEChildren></Outline></Page>"; }
    private static int Main()
    {
        try
        {
            string before = Page("<OE objectID='target'><T>前文 </T><T selected='all'>**测试文字**</T><T> 后文</T></OE><OE objectID='other'><T>保留</T></OE>");
            string after = Page("<OE objectID='target'><T><![CDATA[<span>前文 </span><span style='font-weight:bold'>测试文字</span><span> 后文</span>]]></T></OE><OE objectID='other'><T>保留</T></OE>");
            if (!InlinePostcondition.Matches(before, after)) throw new Exception("Expected fixed replacement");
            if (InlinePostcondition.Matches(before, before)) throw new Exception("Unchanged accepted");
            if (InlinePostcondition.Matches(before, after.Replace("保留", "改变"))) throw new Exception("Outside change accepted");
            if (InlinePostcondition.Matches(before, after.Replace("后文", "丢失"))) throw new Exception("Suffix change accepted");
            if (InlinePostcondition.Matches(before, after.Replace("objectID='target'", "objectID='different'"))) throw new Exception("Different target accepted");
            if(InlinePostcondition.Diagnose(before,before)!="target-text-unchanged") throw new Exception("Unchanged diagnosis");
            if(InlinePostcondition.Diagnose(before,after.Replace("保留","改变"))!="expected-target-text-but-outside-serialization-differs:content-value") throw new Exception("Outside diagnosis");
            if(InlinePostcondition.Diagnose(before,after.Replace("后文","丢失"))!="target-text-unexpected") throw new Exception("Target diagnosis");
            if(InlinePostcondition.Diagnose(before,after.Replace("objectID='target'","objectID='different'"))!="target-identity-not-found") throw new Exception("Identity diagnosis");
            string metadata=after.Replace("<Outline>","<Outline author='fixture'>");
            if(InlinePostcondition.Diagnose(before,metadata)!="expected-target-text-but-outside-serialization-differs:other-attribute") throw new Exception("Metadata must not be silently ignored");
            string reordered=after.Replace("<OE objectID='other'>","<OE selected='none' objectID='other'>");
            if(!InlinePostcondition.Matches(before,reordered)) throw new Exception("Selection-only changes");
            string cosmetic=after.Replace("font-weight:bold","font-weight:normal");
            if(!InlinePostcondition.Matches(before,cosmetic)) throw new Exception("Observer unexpectedly proves visual style");
            string[] originals={"<Size width='1'/>","<OE objectID='a'/>","<OE style='a'/>","<OE><T>x</T></OE>","<OE a='1' b='2'/>"};
            string[] changes={"<Size width='2'/>","<OE objectID='b'/>","<OE style='b'/>","<OE/>","<OE b='2' a='1'/>"};
            string[] categories={"layout-attribute","identity-attribute","style-attribute","child-count","none"};
            for(int i=0;i<originals.Length;i++)
                if(OutsideDiff.Compare(SelectionScreen.Parse(originals[i]),SelectionScreen.Parse(changes[i]))!=categories[i]) throw new Exception("Outside category case "+i);
            string splitBefore=before.Replace("<T>保留</T>","<T>保</T><T>留</T>");
            if(!InlinePostcondition.Diagnose(splitBefore,after).EndsWith("child-count"))throw new Exception("Split export must not match raw shape");
            if(InlinePostcondition.Diagnose(splitBefore,after,before)!="expected-text-and-outside-serialization-match")throw new Exception("Comparable basic baseline failed");
            if(InlinePostcondition.Diagnose(splitBefore,after,before.Replace("前文","变动"))!="baseline-target-text-changed")throw new Exception("Baseline edit accepted");
            if(InlinePostcondition.Diagnose(splitBefore,after.Replace("保留","改变"),before)!="expected-target-text-but-outside-serialization-differs:content-value")throw new Exception("Basic baseline hid outside edit");
            if(!InlinePostcondition.BaselineMatches(before,before))throw new Exception("Identical baseline refused");
            if(InlinePostcondition.BaselineMatches(before,before.Replace("前文","changed")))throw new Exception("Invalid baseline accepted");
            if(!InlinePostcondition.BaselineMatches(splitBefore,before))throw new Exception("Split unrelated text affected baseline");
            Console.WriteLine("PASS 24 fixed inline postcondition/diagnostic cases; text observation is not style or transaction proof");
            return 0;
        }
        catch (Exception ex) { Console.WriteLine("FAIL " + ex.Message); return 1; }
    }
}
