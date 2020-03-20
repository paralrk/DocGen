using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace DocGen.View.Formatters
{
    class PE3SecondPage : A4SecondPage
    {
        public PE3SecondPage(Excel.Worksheet sheet, int firstRow, int pageNumber)
            : base(sheet, firstRow, pageNumber)
        {
        }
        override protected void MergeCells()
        {
            base.MergeCells();
            for (int i = firstRow; i <= firstRow + 29; i++)
            {
                sheet.Range[sheet.Cells[i, 4], sheet.Cells[i, 7]].Merge();
                sheet.Range[sheet.Cells[i, 8], sheet.Cells[i, 20]].Merge();
                sheet.Range[sheet.Cells[i, 21], sheet.Cells[i, 22]].Merge();
                sheet.Range[sheet.Cells[i, 23], sheet.Cells[i, 26]].Merge();
            }
        }

        override protected void DrawBorders()
        {
            base.DrawBorders();

            // specific for PE3
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

            // specific for PE3
            ((Excel.Range)sheet.Cells[firstRow, 3]).Orientation = 90;
            // title default text
            ((Excel.Range)sheet.Cells[firstRow, 3]).Value2 = "Зона";
            sheet.Range[sheet.Cells[firstRow, 4], sheet.Cells[firstRow, 7]].
                Value2 = "Поз. обозначение";
            sheet.Range[sheet.Cells[firstRow, 8], sheet.Cells[firstRow, 20]].
                Value2 = "Наименование";
            sheet.Range[sheet.Cells[firstRow, 21], sheet.Cells[firstRow, 22]].
                Value2 = "Кол.";
            sheet.Range[sheet.Cells[firstRow, 23], sheet.Cells[firstRow, 26]].
                Value2 = "Примечание";

            // font sizes
            ((Excel.Range)sheet.Cells[firstRow, 3]).Font.Size = 14;
            sheet.Range[sheet.Cells[firstRow, 4], sheet.Cells[firstRow, 7]].Font.Size = 11;
            sheet.Range[sheet.Cells[firstRow, 4], sheet.Cells[firstRow, 7]].WrapText = true;
            sheet.Range[sheet.Cells[firstRow, 8], sheet.Cells[firstRow, 20]].Font.Size = 14;
            sheet.Range[sheet.Cells[firstRow, 21], sheet.Cells[firstRow, 22]].Font.Size = 14;
            sheet.Range[sheet.Cells[firstRow, 23], sheet.Cells[firstRow, 26]].Font.Size = 14;
            // text align
            sheet.Range[sheet.Cells[firstRow, 3], sheet.Cells[firstRow, 26]].
                VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            //sheet.Range[sheet.Cells[firstRow + 1, 4], sheet.Cells[firstRow + 29, 26]].
            //    VerticalAlignment = Excel.XlVAlign.xlVAlignBottom;
            //sheet.Range[sheet.Cells[firstRow + 1, 4], sheet.Cells[firstRow + 29, 26]].
            //    HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            //sheet.Range[sheet.Cells[firstRow + 1, 8], sheet.Cells[firstRow + 29, 20]].
            //    HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

            sheet.Range[sheet.Cells[firstRow + 1, 3], sheet.Cells[firstRow + 29, 26]].
                ShrinkToFit = true;
        }
    }
}
