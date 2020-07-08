using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace DocGen.View.Formatters
{
    class D33_UDFirstPage : A4FirstPage
    {
        public D33_UDFirstPage(Excel.Worksheet sheet, int pageCount) : base(sheet, pageCount)
        {
        }

        override protected void SetColumnsWidth()
        {
            // insert extra columns
            Excel.Range range;


            range = sheet.Range["A1:B1"].EntireColumn;
            range.Insert(Excel.XlInsertShiftDirection.xlShiftToRight);
            range = sheet.Range["D1:I1"].EntireColumn;
            range.Insert(Excel.XlInsertShiftDirection.xlShiftToRight);
            range = sheet.Range["K1:M1"].EntireColumn;
            range.Insert(Excel.XlInsertShiftDirection.xlShiftToRight);
            range = sheet.Range["O1:P1"].EntireColumn;
            range.Insert(Excel.XlInsertShiftDirection.xlShiftToRight);
            range = sheet.Range["R1:T1"].EntireColumn;
            range.Insert(Excel.XlInsertShiftDirection.xlShiftToRight);
            range = sheet.Range["V1:Z1"].EntireColumn;
            range.Insert(Excel.XlInsertShiftDirection.xlShiftToRight);


            base.SetColumnsWidth();
        }
        override protected void MergeCells()
        {
            base.MergeCells();

            // specific for D33-UD
            //title
            sheet.Range["C1:I1"].Merge();
            sheet.Range["J1:M1"].Merge();
            sheet.Range["N1:P1"].Merge();
            sheet.Range["Q1:T1"].Merge();
            sheet.Range["U1:Z1"].Merge();
            //rows
            for (int i = 2; i <= 24; i++)
            {
                sheet.Range[sheet.Cells[i, 3], sheet.Cells[i, 9]].Merge();
                sheet.Range[sheet.Cells[i, 10], sheet.Cells[i, 13]].Merge();
                sheet.Range[sheet.Cells[i, 14], sheet.Cells[i, 16]].Merge();
                sheet.Range[sheet.Cells[i, 17], sheet.Cells[i, 20]].Merge();
                sheet.Range[sheet.Cells[i, 21], sheet.Cells[i, 26]].Merge();
            }
        }

        override protected void DrawBorders()
        {
            base.DrawBorders();

            // specific for D33-UD
            // title
            sheet.Range["C1:Z1"].Borders.Weight = Excel.XlBorderWeight.xlMedium;
            // rows
            sheet.Range["C2:Z24"].Borders.Weight = Excel.XlBorderWeight.xlMedium;
            sheet.Range["C2:Z24"].Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle =
                Excel.XlLineStyle.xlLineStyleNone; // внутренние горизонтальные   
        }

        override protected void FillBlank()
        {
            base.FillBlank();

            // specific for D33-UD

            // title default text
            sheet.Range["C1:I1"].Value2 = "Обозначение";
            sheet.Range["J1:M1"].Value2 = "Разработал";
            sheet.Range["N1:P1"].Value2 = "Изготовил";
            sheet.Range["Q1:T1"].Value2 = "Согласовано";
            sheet.Range["U1:Z1"].Value2 = "Утвердил";

        }
    }
}
