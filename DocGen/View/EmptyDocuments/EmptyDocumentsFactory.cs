using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using DocGen.View.Blank;

namespace DocGen.View.EmptyDocuments
{
    class EmptyDocumentsFactory
    {
        public EmptyDocument GetEmptyDocument(Excel.Worksheet sheet)
        {
            string type = ListPage.GetDocumentType(sheet);
            switch (type)
            {
                case "Перечень элементов":
                    return new PE3EmptyDocument();
                case "Спецификация":
                    return new SpecEmptyDocument();
                case "Ведомость покупных изделий":
                    return new VPEmptyDocument();
                case "Д33-УД":
                    return new D33_UDEmptyDocument();
                default:
                    return new PE3EmptyDocument();
            }
        }

        public EmptyDocument GetPE3EmptyDocument()
        {
            return new PE3EmptyDocument();
        }

        public EmptyDocument GetSpecificationEmptyDocument()
        {
            return new SpecEmptyDocument();
        }

        public EmptyDocument GetVPEmptyDocument()
        {
            return new VPEmptyDocument();
        }

        public EmptyDocument GetD33_UDEmptyDocument()
        {
            return new D33_UDEmptyDocument();
        }
    }
}
