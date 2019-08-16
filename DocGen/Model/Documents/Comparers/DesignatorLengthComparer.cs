using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocGen.Model;

namespace DocGen.Model.Documents.Comparers
{
    class DesignatorLengthComparer : IComparer<Components>
    {
        public int Compare(Components c1, Components c2)
        {
            if (c1.GetDesignators().Length > c2.GetDesignators().Length)
            {
                return 1;
            } else if (c1.GetDesignators().Length == c2.GetDesignators().Length)
            {
                return 0;
            } else  // if (c1.Designator.Length < c2.Designator.Length)
            {
                return -1;
            }
        }
    }
}
