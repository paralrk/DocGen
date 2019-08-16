using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace DocGen.View.Formatters
{
    class SpecFirstPage : A4FirstPage
    {
        public SpecFirstPage(Excel.Worksheet sheet, int pageCount) : base(sheet, pageCount)
        {
        }

        override protected void SetColumnsWidth()
        {
            // insert extra columns
            Excel.Range range;


            range = sheet.Range["A1:B1"].EntireColumn;
            range.Insert(Excel.XlInsertShiftDirection.xlShiftToRight);
            range = sheet.Range["F1:F1"].EntireColumn;
            range.Insert(Excel.XlInsertShiftDirection.xlShiftToRight);
            range = sheet.Range["H1:O1"].EntireColumn;
            range.Insert(Excel.XlInsertShiftDirection.xlShiftToRight);
            range = sheet.Range["Q1:U1"].EntireColumn;
            range.Insert(Excel.XlInsertShiftDirection.xlShiftToRight);
            range = sheet.Range["W1:W1"].EntireColumn;
            range.Insert(Excel.XlInsertShiftDirection.xlShiftToRight);
            range = sheet.Range["Y1:Z1"].EntireColumn;
            range.Insert(Excel.XlInsertShiftDirection.xlShiftToRight);


            base.SetColumnsWidth();
        }
        override protected void MergeCells()
        {
            base.MergeCells();

            // specific for Specification
            //title
            sheet.Range["E1:F1"].Merge();
            sheet.Range["G1:O1"].Merge();
            sheet.Range["P1:U1"].Merge();
            sheet.Range["V1:W1"].Merge();
            sheet.Range["X1:Z1"].Merge();
            //rows
            for (int i = 2; i <= 24; i++)
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
            sheet.Range["C1:Z1"].Borders.Weight = Excel.XlBorderWeight.xlMedium;
            // rows
            sheet.Range["C2:Z24"].Borders.Weight = Excel.XlBorderWeight.xlMedium;
            sheet.Range["C2:Z24"].Borders[Excel.XlBordersIndex.xlInsideHorizontal].Weight =
                Excel.XlBorderWeight.xlThin; // внутренние горизонтальные            
        }

        override protected void FillBlank()
        {
            base.FillBlank();

            // specific for Specification
            sheet.Range["C1:F1"].Orientation = 90;
            sheet.Range["V1:W1"].Orientation = 90;
            // title default text
            sheet.Range["C1"].Value2 = "Формат";
            sheet.Range["D1"].Value2 = "Зона";
            sheet.Range["E1:F1"].Value2 = "Поз.";
            sheet.Range["G1:O1"].Value2 = "Обозначение";
            sheet.Range["P1:U1"].Value2 = "Наименование";
            sheet.Range["V1:W1"].Value2 = "Кол.";
            sheet.Range["X1:Z1"].Value2 = "Приме-чание";
            sheet.Range["M36:R36"].Value2 = "";
            // font sizes
            sheet.Range["C1"].Font.Size = 10;
            sheet.Range["D1:Z1"].Font.Size = 14;
            sheet.Range["X1:Z1"].WrapText = true;
            sheet.Range["M36:R36"].Font.Size = 12;
            // text align
            sheet.Range["C1:Z1"].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            sheet.Range["C2:Z24"].VerticalAlignment = Excel.XlVAlign.xlVAlignBottom;
            sheet.Range["C1:Z24"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            sheet.Range["G2:O24"].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            sheet.Range["P2:U24"].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

            sheet.Range["E2:Z24"].ShrinkToFit = true;
        }
    }
}
