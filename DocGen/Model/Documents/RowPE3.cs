using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocGen.Model.Documents
{
    class RowPE3 : IRow
    {
        public string Zone { get; set; }
        public string Designator { get; set; }
        public string Name { get; set; }
        public int Quantity { get; set; }
        public string Note { get; set; }

        // constructors

        public RowPE3()
        {
            this.Zone = "";
            this.Designator = "";
            this.Name = "";
            this.Quantity = 0;
            this.Note = "";
        }

        public RowPE3(string zone, string designator, string name, int quantity, string note)
        {
            this.Zone = zone;
            this.Designator = designator;
            this.Name = name;
            this.Quantity = quantity;
            this.Note = note;
        }
    }
}
