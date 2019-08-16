using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace DocGen.Utils
{
    static class DesignatorShortener
    {
        public static string Short(LinkedList<string> designators)
        {

            if (designators.Count < 3)
            {
                return String.Join(", ", designators);
            }

            string shortDes = "";
            Regex regexDigitals = new Regex(@"[0-9]+");
            Match match = null;
            Regex regexLetters = new Regex(@"[a-zA-Z]+");
            Match matchLetters = regexLetters.Match(designators.First());
            string baseDes = matchLetters.Groups[0].Value;

            // variables to find sequence in designators
            // if diff > 2, designators are combined
            // Example: DA1, DA2, DA3, DA4, DA6, DA7 -> DA1...DA4, DA6, DA7
            int current = 0;
            int sequence = 0;
            int diff = 0;
            match = regexDigitals.Match(designators.First());
            string mathced = match.Groups[0].Value;
            int last = Convert.ToInt32(mathced);
            shortDes = designators.First();

            for (int i = 1; i < designators.Count; i++)
            {
                match = regexDigitals.Match(designators.ElementAt(i));
                mathced = match.Groups[0].Value;
                current = Convert.ToInt32(mathced);
                diff = current - last;

                if (diff == 1)
                {
                    sequence++;
                    //last = current;
                }

                // if sequence is broken
                // short designators (DA1...DA3) or just add (DA3, DA5)
                if (diff > 1)
                {
                    if (sequence >= 2)
                    {
                        shortDes += "..." + baseDes + last + ", " + baseDes + current;
                    }
                    if (sequence == 1)
                    {
                        shortDes += ", " + baseDes + last + ", " + baseDes + current;
                    }
                    if (sequence == 0)
                    {
                        shortDes += ", " + baseDes + current;
                    }
                    sequence = 0;
                }
                last = current;
            } // for

            if (sequence >= 2)
            {
                shortDes += "..." + baseDes + last;
            }
            if (sequence == 1)
            {
                shortDes += ", " + baseDes + last;
            }

            return shortDes;
        }
    }
}
