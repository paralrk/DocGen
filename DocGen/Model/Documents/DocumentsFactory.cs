using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using DocGen.View.Blank;

namespace DocGen.Model.Documents
{
    class DocumentsFactory
    {
        public Document GetPE3Document()
        {
            return new PE3();
        }

        public Document GetSpecificationDocument()
        {
            return new Specification();
        }
    }
}
