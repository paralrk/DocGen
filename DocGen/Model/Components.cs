using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocGen.Model
{
    class Components
    {
        private LinkedList<string> designators = new LinkedList<string>();
        public int Quantity { get; set; } = 0;
        public Part Part { get; set; }
        public int NoteNumber{ get; set; } = 0;
        public int NoteNumber1 { get; set; } = 0;

        public Components ()
        {
            this.Part = new Part();
        }

        public void AddComponent(Components c)
        {
            designators.AddLast(c.GetDesignators());
            Quantity += c.Quantity;
            Part = c.Part;
        }

        public string GetDesignators()
        {
            if (designators.Count == 0)
            {
                return "";
            }
            if (designators.Count > 1)
            {
                return String.Join(", ", designators);
            } else
            {
                return designators.First(); // if Count == 1
            }
        }

        public LinkedList<string> GetDesignatorsList()
        {
            return designators;
        }

        public void AddDesignator (string des)
        {
            designators.AddLast(des);
        }

        public bool IsEqualHeader(Components c)
        {
            return (this.Part.Type.Equals(c.Part.Type) &&
                this.Part.Manufacturer.Equals(c.Part.Manufacturer));
        }
    }
}
