using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

using DocGen.Utils;

namespace DocGen.View.Blank
{
    class A4SecondPageBlankTextFiller
    {
        public void Fill(Excel.Worksheet sheet, int firstRow)
        {
            Excel.Worksheet list = ListPage.GetListPage();
            string column = ListPage.GetDocumentColumn(sheet.Name);
            // for example: Value2 = "=Содержание!$B$3"
            sheet.Range[sheet.Cells[firstRow + 30, 13], sheet.Cells[firstRow + 33, 25]].
                Value2 = "=Содержание!$" + column + "$" + (int)Names.Designation; // Обозначение
            sheet.Range[sheet.Cells[firstRow + 29, 2], sheet.Cells[firstRow + 33, 2]].
                Value2 = "=Содержание!$" + column + "$" + (int)Names.OriginalInvNumber; // Инв№Подл
            sheet.Range[sheet.Cells[firstRow + 25, 2], sheet.Cells[firstRow + 28, 2]].
                Value2 = "=Содержание!$" + column + "$" + (int)Names.SignDateOriginal; // ПодпИДатаПодл
            sheet.Range[sheet.Cells[firstRow + 22, 2], sheet.Cells[firstRow + 24, 2]].
                Value2 = "=Содержание!$" + column + "$" + (int)Names.InsteadInvNumber; // ВзамИнв№
            sheet.Range[sheet.Cells[firstRow + 19, 2], sheet.Cells[firstRow + 21, 2]].
                Value2 = "=Содержание!$" + column + "$" + (int)Names.DublicateInvNumber; //  Инв№Дубл		
            sheet.Range[sheet.Cells[firstRow + 15, 2], sheet.Cells[firstRow + 18, 2]].
                Value2 = "=Содержание!$" + column + "$" + (int)Names.SignDateDublicate; //  ПодпИДатаДубл	

        }
    }
}
