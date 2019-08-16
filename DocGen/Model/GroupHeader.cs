using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocGen.Model
{
    class GroupHeader : IEquatable<GroupHeader>
    {
        public string Manufacturer { get; set; }
        public string TypeDescription { get; set; }
        public int NoteNumber { get; set; }
        public int NoteNumber1 { get; set; }
        public int Count { get; set; }

        public override bool Equals(Object obj)
        {
            //Check for null and compare run-time types.
            if ((obj == null) || !this.GetType().Equals(obj.GetType()))
            {
                return false;
            }
            else
            {
                GroupHeader header = (GroupHeader)obj;
                return Equals(header);
            }
        }

        public bool Equals(GroupHeader header)
        {
            return (this.Manufacturer.Equals(header.Manufacturer) &&
                this.TypeDescription.Equals(header.TypeDescription));
        }

        public override int GetHashCode()
        {
            int hash = 0;
            hash = this.Manufacturer.GetHashCode() +
                this.TypeDescription.GetHashCode();
            return hash;
        }
    }
}
