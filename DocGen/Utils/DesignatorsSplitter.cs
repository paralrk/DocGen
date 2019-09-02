using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace DocGen.Utils
{
    static class DesignatorsSplitter
    {
        public static string[] SplitDesignators(string designators)
        {
            // ignore "," and " "
            // example 
            // "C1, C2, C10, C12" (1 stirng) -> "C1", "C2", "C10", "C12" (4 stirngs)
            string regPattern = @"[^,^ ]+";
            String pattern = regPattern;
            MatchCollection matchList = Regex.Matches(designators, pattern);
            string[] splitted = matchList.Cast<Match>().
                                Select(match => match.Value).ToArray();
            return splitted.OrderBy(s => s.Length).ThenBy(s => s).ToArray();
        }
    }
}
