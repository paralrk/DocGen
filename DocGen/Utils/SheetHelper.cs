using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Interop.Excel;

namespace DocGen.Utils
{
    class SheetHelper
    {

        public static Excel.Worksheet getSheet()
        {
            return (Excel.Worksheet)Globals.ThisAddIn.Application.Worksheets.Add();
        }

        public static Excel.Worksheet getSheet(string name)
        {

            if (isExisting(name))
            {
                return (Excel.Worksheet)Globals.ThisAddIn.Application.Worksheets[name];
                //return (Excel.Worksheet)workbook.Worksheets[name];
            }
            else
            {
                Excel.Worksheet sheet = (Excel.Worksheet)Globals.ThisAddIn.Application.Worksheets.Add();
                sheet.Name = name;
                return sheet;
            }
        }

        public static bool isExisting(string name)
        {
            Excel.Workbook workbook = (Excel.Workbook)Globals.ThisAddIn.Application.ActiveWorkbook;
            bool found = false;
            foreach (Excel.Worksheet sheet in workbook.Sheets)
            {
                if (sheet.Name == name)
                {
                    found = true;
                    break;
                }
            }
            return found;
        }
    }
}
