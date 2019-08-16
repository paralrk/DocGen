using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocGen.Model.Documents.Templates
{
    class PartsSection
    {
        private List<RowSpec> parts;

        public PartsSection()
        {
            parts = new List<RowSpec>();
            parts.Add(new RowSpec("", "", "", "", "Детали", 0, ""));
            parts.Add(new RowSpec());
            parts.Add(new RowSpec("", "", "1", "", "Плата печатная", 1, ""));
        }

        public List<RowSpec> GetDocuments()
        {
            return parts;
        }
    }
}
