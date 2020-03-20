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
        public IDocument GetPE3Document()
        {
            return new PE3();
        }

        public IDocument GetSpecificationDocument()
        {
            return new Specification();
        }

        public IDocument GetSWSpecificationDocument()
        {
            return new SWSpecification();
        }
    }
}
