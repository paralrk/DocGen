using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocGen.Model
{
    class Component
    {
        public string Designator { get; set; }
        public int Quantity { get; set; }
        public Part Part {  get; set; }


    }

}
