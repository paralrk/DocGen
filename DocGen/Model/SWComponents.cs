using DocGen.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocGen.Model
{
    class SWComponents
    {
		public String Format { get; set; } = "";
		public String Position { get; set; } = "";
		public String Designation { get; set; } = "";
		public String Name { get; set; } = "";
		public int Quantity { get; set; } = 0;
		public String Note { get; set; } = "";
		public String DocumentSection { get; set; } = "";
		public String Class { get; set; } = "";
		public String Gost { get; set; } = "";
		public String SizesParametres { get; set; } = "";
		public String Replacement { get; set; } = "";
		public int ReplacementNumber { get; set; } = 0;
    }
}
