using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using DocGen.View.Blank;

namespace DocGen.View.Formatters
{
    class FormattersFactory
    {
        public Formatter GetFormatter()
        {
            Excel.Worksheet sheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            string type = ListPage.GetDocumentType(sheet);
            switch (type)
            {
                case "Перечень элементов":
                    return new PE3Formatter();
                case "Спецификация":
                    return new SpecFormatter();
                case "Ведомость покупных изделий":
                    return new VPFormatter();
                case "Д33-УД":
                    return new D33_UDFormatter();
                default:
                    return null;
            }
        }
    }
}
