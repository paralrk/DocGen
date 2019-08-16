using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocGen.Model.Documents
{
    class RowSpec : IRow
    {
        public string Format { get; set; }
        public string Zone { get; set; }
        public string Position { get; set; }
        public string Designation { get; set; }
        public string Name { get; set; }
        public int Quantity { get; set; }
        public string Note { get; set; }

        // constructors

        public RowSpec()
        {
            this.Format = "";
            this.Zone = "";
            this.Position = "";
            this.Designation = "";
            this.Name = "";
            this.Quantity = 0;
            this.Note = "";
        }

        public RowSpec(string format,
                        string zone,
                        string position,
                        string designation,
                        string name, 
                        int quantity, 
                        string note)
        {
            this.Format = format;
            this.Zone = zone;
            this.Position = position;
            this.Designation = designation;
            this.Name = name;
            this.Quantity = quantity;
            this.Note = note;
        }
    }
}
