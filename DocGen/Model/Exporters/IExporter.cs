using DocGen.Model.Documents;
using DocGen.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocGen.Model.Exporters
{
    interface IExporter
    {
        void Export(Document doc);
    }
}
