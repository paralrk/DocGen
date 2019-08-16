using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocGen.Model
{
    class Part
    {
        public Part()
        {
            this.Type = "";
            this.ManufacturerPartNumber = "";
            this.Description = "";
            this.Manufacturer = "";
            this.Note = "";
            this.Note1 = "";
        }
        public string Type { get; set; }
        public string ManufacturerPartNumber { get; set; }
        public string Description { get; set; }
        public string Manufacturer { get; set; }
        public string Note { get; set; }
        public string Note1 { get; set; }

    }
}
