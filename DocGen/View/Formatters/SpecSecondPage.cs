using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace DocGen.View.Formatters
{
    class SpecSecondPage : A4SecondPage
    {
        public SpecSecondPage(Excel.Worksheet sheet, int firstRow, int pageNumber)
            : base(sheet, firstRow, pageNumber)
        {
        }

        override protected void MergeCells()
        {
            base.MergeCells();
            for (int i = firstRow; i <= firstRow + 29; i++)
            {
                sheet.Range[sheet.Cells[i, 5], sheet.Cells[i, 6]].Merge();
                sheet.Range[sheet.Cells[i, 7], sheet.Cells[i, 15]].Merge();
                sheet.Range[sheet.Cells[i, 16], sheet.Cells[i, 21]].Merge();
                sheet.Range[sheet.Cells[i, 22], sheet.Cells[i, 23]].Merge();
                sheet.Range[sheet.Cells[i, 24], sheet.Cells[i, 26]].Merge();
            }
        }

        override protected void DrawBorders()
        {
            base.DrawBorders();

            // specific for Specification
            // title
            sheet.Range[sheet.Cells[firstRow, 3], sheet.Cells[firstRow, 26]].
                Borders.Weight = Excel.XlBorderWeight.xlMedium;
            // rows
            sheet.Range[sheet.Cells[firstRow + 1, 3], sheet.Cells[firstRow + 29, 26]].
                Borders.Weight = Excel.XlBorderWeight.xlMedium;
            sheet.Range[sheet.Cells[firstRow + 1, 3], sheet.Cells[firstRow + 29, 26]].
                Borders[Excel.XlBordersIndex.xlInsideHorizontal].Weight = Excel.XlBorderWeight.xlThin; // внутренние горизонтальные            
        }

        override protected void FillBlank()
        {
            base.FillBlank();

            // specific for Specification
            sheet.Range[sheet.Cells[firstRow, 3], sheet.Cells[firstRow, 6]].
                Orientation = 90;
            sheet.Range[sheet.Cells[firstRow, 22], sheet.Cells[firstRow, 23]].
                Orientation = 90;
            // title default text
            ((Excel.Range)sheet.Cells[firstRow, 3]).Value2 = "Формат";
            ((Excel.Range)sheet.Cells[firstRow, 4]).Value2 = "Зона";
            sheet.Range[sheet.Cells[firstRow, 5], sheet.Cells[firstRow, 6]].
                Value2 = "Поз.";
            sheet.Range[sheet.Cells[firstRow, 7], sheet.Cells[firstRow, 15]].
                Value2 = "Обозначение";
            sheet.Range[sheet.Cells[firstRow, 16], sheet.Cells[firstRow, 21]].
                Value2 = "Наименование";
            sheet.Range[sheet.Cells[firstRow, 22], sheet.Cells[firstRow, 23]].
                Value2 = "Кол.";
            sheet.Range[sheet.Cells[firstRow, 24], sheet.Cells[firstRow, 26]].
                Value2 = "Приме-чание";

            // font sizes
            ((Excel.Range)sheet.Cells[firstRow, 3]).Font.Size = 10;
            sheet.Range[sheet.Cells[firstRow, 4], sheet.Cells[firstRow, 26]].Font.Size = 14;
            sheet.Range[sheet.Cells[firstRow, 24], sheet.Cells[firstRow, 26]].WrapText = true;

            // text align
            sheet.Range[sheet.Cells[firstRow, 3], sheet.Cells[firstRow, 26]].
                VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            sheet.Range[sheet.Cells[firstRow, 3], sheet.Cells[firstRow, 26]].
                HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            sheet.Range[sheet.Cells[firstRow + 1, 4], sheet.Cells[firstRow + 29, 26]].
                VerticalAlignment = Excel.XlVAlign.xlVAlignBottom;
            sheet.Range[sheet.Cells[firstRow + 1, 4], sheet.Cells[firstRow + 29, 26]].
                HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            sheet.Range[sheet.Cells[firstRow + 1, 7], sheet.Cells[firstRow + 29, 15]].
                HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            sheet.Range[sheet.Cells[firstRow + 1, 16], sheet.Cells[firstRow + 29, 21]].
                HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

            sheet.Range[sheet.Cells[firstRow + 1, 3], sheet.Cells[firstRow + 29, 26]].
                ShrinkToFit = true;
        }
    }
}
