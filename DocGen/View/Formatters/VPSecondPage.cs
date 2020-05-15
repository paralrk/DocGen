using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace DocGen.View.Formatters
{
    class VPSecondPage : A3SecondPage
    {
        public VPSecondPage(Excel.Worksheet sheet, int firstRow, int pageNumber)
            : base(sheet, firstRow, pageNumber)
        {
        }

        override protected void MergeCells()
        {
            base.MergeCells();

            //title
            sheet.Range[sheet.Cells[firstRow, 3], sheet.Cells[firstRow + 1, 3]].Merge(); // С
            sheet.Range[sheet.Cells[firstRow, 4], sheet.Cells[firstRow + 1, 14]].Merge(); // D-N
            sheet.Range[sheet.Cells[firstRow, 15], sheet.Cells[firstRow + 1, 17]].Merge(); // O-Q
            sheet.Range[sheet.Cells[firstRow, 18], sheet.Cells[firstRow + 1, 26]].Merge(); // R-Z
            sheet.Range[sheet.Cells[firstRow, 27], sheet.Cells[firstRow + 1, 30]].Merge(); // AA-AD
            sheet.Range[sheet.Cells[firstRow, 31], sheet.Cells[firstRow + 1, 36]].Merge(); // AE-AJ
            sheet.Range[sheet.Cells[firstRow, 37], sheet.Cells[firstRow, 43]].Merge(); // Количество AK-AQ
            sheet.Range[sheet.Cells[firstRow + 1, 39], sheet.Cells[firstRow + 1, 41]].Merge(); // AM-AO
            sheet.Range[sheet.Cells[firstRow + 1, 42], sheet.Cells[firstRow + 1, 43]].Merge(); // AP-AQ
            sheet.Range[sheet.Cells[firstRow, 44], sheet.Cells[firstRow + 1, 46]].Merge(); // AR-AT


            for (int i = firstRow + 2; i <= firstRow + 29; i++)
            {
                sheet.Range[sheet.Cells[i, 4], sheet.Cells[i, 14]].Merge();
                sheet.Range[sheet.Cells[i, 15], sheet.Cells[i, 17]].Merge();
                sheet.Range[sheet.Cells[i, 18], sheet.Cells[i, 26]].Merge();
                sheet.Range[sheet.Cells[i, 27], sheet.Cells[i, 30]].Merge();
                sheet.Range[sheet.Cells[i, 31], sheet.Cells[i, 36]].Merge();
                sheet.Range[sheet.Cells[i, 39], sheet.Cells[i, 41]].Merge();
                sheet.Range[sheet.Cells[i, 42], sheet.Cells[i, 43]].Merge();
                sheet.Range[sheet.Cells[i, 44], sheet.Cells[i, 46]].Merge();
            }
        }

        override protected void DrawBorders()
        {
            base.DrawBorders();

            // specific for VP
            // title
            sheet.Range[sheet.Cells[firstRow, 3], sheet.Cells[firstRow + 1, 46]].
                Borders.Weight = Excel.XlBorderWeight.xlMedium;
            // rows
            sheet.Range[sheet.Cells[firstRow + 2, 3], sheet.Cells[firstRow + 29, 46]].
                Borders.Weight = Excel.XlBorderWeight.xlMedium;
            sheet.Range[sheet.Cells[firstRow + 2, 3], sheet.Cells[firstRow + 29, 46]].
                Borders[Excel.XlBordersIndex.xlInsideHorizontal].Weight = Excel.XlBorderWeight.xlThin; // внутренние горизонтальные            
        }

        override protected void FillBlank()
        {
            base.FillBlank();

            // specific for VP
            sheet.Range[sheet.Cells[firstRow, 3], sheet.Cells[firstRow + 1, 3]].
                Orientation = 90;
            // title default text

            sheet.Range[sheet.Cells[firstRow, 3], sheet.Cells[firstRow + 1, 3]].Value2 = "№ строки"; // С
            sheet.Range[sheet.Cells[firstRow, 4], sheet.Cells[firstRow + 1, 14]].Value2 = "Наименование"; // D-N
            sheet.Range[sheet.Cells[firstRow, 15], sheet.Cells[firstRow + 1, 17]].Value2 = "Код продукции"; // O-Q
            sheet.Range[sheet.Cells[firstRow, 18], sheet.Cells[firstRow + 1, 26]].Value2 = "Обозначение документа";  // R-Z
            sheet.Range[sheet.Cells[firstRow, 27], sheet.Cells[firstRow + 1, 30]].Value2 = "Поставщик";  // AA-AD
            sheet.Range[sheet.Cells[firstRow, 31], sheet.Cells[firstRow + 1, 43]].Value2 = "Куда входит (обозначение)";  // AE-AJ
            sheet.Range[sheet.Cells[firstRow, 37], sheet.Cells[firstRow, 43]].Value2 = "Количество";  // Количество AK-AQ
            ((Excel.Range)sheet.Cells[firstRow + 1, 37]).Value2 = "на из- делие";  // AK-AK
            ((Excel.Range)sheet.Cells[firstRow + 1, 38]).Value2 = "в ком- плекты";  // AL-AL
            sheet.Range[sheet.Cells[firstRow + 1, 39], sheet.Cells[firstRow + 1, 41]].Value2 = "на ре- гулир.";  // AM-AO
            sheet.Range[sheet.Cells[firstRow + 1, 42], sheet.Cells[firstRow + 1, 43]].Value2 = "всего";  // AP-AQ
            sheet.Range[sheet.Cells[firstRow, 44], sheet.Cells[firstRow + 1, 46]].Value2 = "Приме- чание";  // AR-AT


            for (int i = firstRow + 2, line = 1; i <= firstRow + 29; i++, line++)
            {
                ((Excel.Range)sheet.Cells[i, 3]).Value2 = line;
            }


            // font sizes
            sheet.Range[sheet.Cells[firstRow, 3], sheet.Cells[firstRow + 1, 3]].Font.Size = 11;
            sheet.Range[sheet.Cells[firstRow, 4], sheet.Cells[firstRow + 1, 46]].Font.Size = 14;
            sheet.Range[sheet.Cells[firstRow + 1, 37], sheet.Cells[firstRow + 1, 43]].Font.Size = 11;
            sheet.Range[sheet.Cells[firstRow, 4], sheet.Cells[firstRow + 1, 46]].WrapText = true;

            // text align
            sheet.Range[sheet.Cells[firstRow, 3], sheet.Cells[firstRow + 1, 46]]
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            sheet.Range[sheet.Cells[firstRow, 3], sheet.Cells[firstRow + 1, 46]]
                .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            sheet.Range[sheet.Cells[firstRow + 2, 3], sheet.Cells[firstRow + 29, 46]].ShrinkToFit = true;
        }
    }
}
