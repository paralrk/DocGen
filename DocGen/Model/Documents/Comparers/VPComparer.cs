using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace DocGen.Model.Documents.Comparers
{
    class VPComparer : IComparer<Components>
    {
        Dictionary<char, int> map = new Dictionary<char, int>()
        {
            { 'R', 26},
            { 'C', 25},
            { 'V', 24},
            { 'U', 23},
            { 'D', 22},
            { 'K', 21},
            { 'S', 20},
            { 'Y', 19},
            { 'F', 18},
            { 'X', 17},
            { 'T', 16},
            { 'L', 15},
            { 'Z', 14},
            { 'M', 13},
            { 'P', 12},
            { 'A', 11},
            { 'B', 10},
            { 'H', 9},
            { 'E', 8},
            { 'G', 7},
            { 'Q', 6},
            { 'W', 5},
            { 'I', 4},
            { 'J', 3},
            { 'N', 2},
            { 'O', 1}
        };

        public int Compare(Components c1, Components c2)
        {
            string des1 = c1.GetDesignators();
            string des2 = c2.GetDesignators();

            Regex regex = new Regex(@"[0-9]", RegexOptions.Compiled);

            des1 = regex.Replace(des1, "");
            des2 = regex.Replace(des2, "");

            if (String.IsNullOrEmpty(des1) || !map.ContainsKey(des1[0]))
            {
                return -1;
            }
            if (String.IsNullOrEmpty(des2) || !map.ContainsKey(des2[0]))
            {
                return 1;
            }

            if (des1.Equals(des2))
            {
                return 0;
            }

            if (map[des1[0]] == map[des2[0]])
            {

                if (des1.Length > 1 && des2.Length > 1)
                {
                    return des1[1].CompareTo(des2[1]);
                }

            }

            if (map[des1[0]] < map[des2[0]])
            {
                return 1;
            }
            else
            {
                return -1;
            }
        }
    }
}
