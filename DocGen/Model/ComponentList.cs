using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocGen.Model
{
    class ComponentList : IEnumerable<Group>
    {
        public List<Group> Groups { get; set; }

        public IEnumerator<Group> GetEnumerator()
        {
            return Groups.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }

    }
}
