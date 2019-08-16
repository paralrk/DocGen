using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocGen.Model;

namespace DocGen.Model.Documents.Comparers
{
    class DesignatorComparer : IComparer<Components>
    {
        public int Compare(Components c1, Components c2)
        {
            return c1.GetDesignators().CompareTo(c2.GetDesignators());
        }
    }
}
