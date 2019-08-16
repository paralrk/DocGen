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
    class StyleHelper
    {
        public static Excel.Style getDocGenMainStyle()
        {
            string styleName = "DocGenMainStyle";
            Excel.Workbook workbook = (Excel.Workbook)Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Style style = null;
            if (isStyleExists(styleName))
            {
                style = workbook.Styles[styleName];
            }
            else
            {
                style = workbook.Styles.Add(styleName);
            }

            style.Font.Name = "Isocpeur";
            style.Font.Size = 14;
            style.Font.Italic = true;
            return style;

        }

        private static bool isStyleExists(string name)
        {
            Excel.Workbook workbook = (Excel.Workbook)Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Styles styles = workbook.Styles;

            bool found = false;
            for (int i = 1; i <= styles.Count; i++)
            {
                if (styles[i].Name == name)
                {
                    found = true;
                    break;
                }
            }
            return found;
        }
    }
}
