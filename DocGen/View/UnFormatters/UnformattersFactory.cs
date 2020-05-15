using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using DocGen.View.Blank;

namespace DocGen.View.Unformatters
{
    class UnformattersFactory
    {
        public Unformatter GetUnformatter()
        {
            Excel.Worksheet sheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            string type = ListPage.GetDocumentType(sheet);
            switch (type)
            {
                case "Перечень элементов":
                    return new PE3Unformatter();
                case "Спецификация":
                    return new SpecUnformatter();
                case "Ведомость покупных изделий":
                    return new VPUnformatter();
                default:
                    return null;
            }
        }
    }
}
