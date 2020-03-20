using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace DocGen.View.Formatters
{
    class PE3FirstPage : A4FirstPage
    {

        public PE3FirstPage(Excel.Worksheet sheet, int pageCount) : base(sheet, pageCount)
        {
        }

        override protected void SetColumnsWidth()
        {
            // insert extra columns
            Excel.Range column;
            column = (Excel.Range)sheet.Columns[1];
            for (int i = 1; i <= 2; i++)
            {
                column.Insert();
            }
            column = (Excel.Range)sheet.Columns[5];
            for (int i = 5; i <= 7; i++)
            {
                column.Insert();
            }
            column = (Excel.Range)sheet.Columns[9];
            for (int i = 9; i <= 20; i++)
            {
                column.Insert();
            }
            column = (Excel.Range)sheet.Columns[22];
            column.Insert();
            column = (Excel.Range)sheet.Columns[24];
            for (int i = 24; i <= 26; i++)
            {
                column.Insert();
            }

            base.SetColumnsWidth();
        }

        override protected void MergeCells()
        {
            base.MergeCells();

            // specific for PE3
            //title
            sheet.Range["D1:G1"].Merge();
            sheet.Range["H1:T1"].Merge();
            sheet.Range["U1:V1"].Merge();
            sheet.Range["W1:Z1"].Merge();
            //rows
            for (int i = 2; i <= 24; i++)
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
            sheet.Range["C1:Z1"].Borders.Weight = Excel.XlBorderWeight.xlMedium;
            // rows
            sheet.Range["C2:Z24"].Borders.Weight = Excel.XlBorderWeight.xlMedium;
            sheet.Range["C2:Z24"].Borders[Excel.XlBordersIndex.xlInsideHorizontal].Weight =
                Excel.XlBorderWeight.xlThin; // внутренние горизонтальные            
        }
        override protected void FillBlank()
        {
            base.FillBlank();

            // specific for PE3
            sheet.Range["C1"].Orientation = 90;
            // title default text
            sheet.Range["C1"].Value2 = "Зона";
            sheet.Range["D1:G1"].Value2 = "Поз. обозначение";
            sheet.Range["H1:T1"].Value2 = "Наименование";
            sheet.Range["U1:V1"].Value2 = "Кол.";
            sheet.Range["W1:Z1"].Value2 = "Примечание";
            sheet.Range["M36:R36"].Value2 = "Перечень элементов";
            // font sizes
            sheet.Range["C1"].Font.Size = 14;
            sheet.Range["D1:G1"].Font.Size = 11;
            sheet.Range["D1:G1"].WrapText = true;
            sheet.Range["H1:T1"].Font.Size = 14;
            sheet.Range["U1:V1"].Font.Size = 14;
            sheet.Range["W1:Z1"].Font.Size = 14;
            sheet.Range["M36:R36"].Font.Size = 12;
            // text align
            sheet.Range["C1:Z1"].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            // sheet.Range["C2:Z24"].VerticalAlignment = Excel.XlVAlign.xlVAlignBottom;
            sheet.Range["C1:Z24"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            // sheet.Range["H2:T24"].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

            sheet.Range["C2:Z24"].ShrinkToFit = true;
        }
    }
}
