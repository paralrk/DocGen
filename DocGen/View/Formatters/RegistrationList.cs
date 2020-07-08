using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace DocGen.View.Formatters
{
    class RegistrationList : A4SecondPage
    {
        public RegistrationList(Excel.Worksheet sheet, int firstRow, int pageNumber)
            : base(sheet, firstRow, pageNumber)
        {
        }

        override protected void SetRowsHeight()
        {
            base.SetRowsHeight();

            string str;
            Excel.Range range;
            str = firstRow + ":" + (firstRow + 2);
            range = sheet.Range[str];
            range.RowHeight = 29;
            //str = (firstRow + 1) + ":" + (firstRow + 1);
            //range = sheet.Range[str];
            //range.RowHeight = 25;
            //str = (firstRow + 2) + ":" + (firstRow + 2);
            //range = sheet.Range[str];
            //range.RowHeight = 25;
        }

        override protected void MergeCells()
        {
            base.MergeCells();
            // title
            sheet.Range[sheet.Cells[firstRow, 3], sheet.Cells[firstRow, 26]].Merge();
            // изм.
            sheet.Range[sheet.Cells[firstRow + 1, 3], sheet.Cells[firstRow + 2, 4]].Merge();
            // номера листов (страниц) и графы под
            sheet.Range[sheet.Cells[firstRow + 1, 5], sheet.Cells[firstRow + 1, 14]].Merge();
            sheet.Range[sheet.Cells[firstRow + 2, 5], sheet.Cells[firstRow + 2, 7]].Merge();
            sheet.Range[sheet.Cells[firstRow + 2, 8], sheet.Cells[firstRow + 2, 9]].Merge();
            sheet.Range[sheet.Cells[firstRow + 2, 10], sheet.Cells[firstRow + 2, 12]].Merge();
            sheet.Range[sheet.Cells[firstRow + 2, 13], sheet.Cells[firstRow + 2, 14]].Merge();
            // всего
            sheet.Range[sheet.Cells[firstRow + 1, 15], sheet.Cells[firstRow + 2, 16]].Merge();
            // № докум
            sheet.Range[sheet.Cells[firstRow + 1, 17], sheet.Cells[firstRow + 2, 17]].Merge();
            // входящий №
            sheet.Range[sheet.Cells[firstRow + 1, 18], sheet.Cells[firstRow + 2, 21]].Merge();
            // подп
            sheet.Range[sheet.Cells[firstRow + 1, 22], sheet.Cells[firstRow + 2, 24]].Merge();
            // дата
            sheet.Range[sheet.Cells[firstRow + 1, 25], sheet.Cells[firstRow + 2, 26]].Merge();

            for (int i = firstRow + 3; i <= firstRow + 29; i++)
            {
                sheet.Range[sheet.Cells[i, 3], sheet.Cells[i, 4]].Merge();
                sheet.Range[sheet.Cells[i, 5], sheet.Cells[i, 7]].Merge();
                sheet.Range[sheet.Cells[i, 8], sheet.Cells[i, 9]].Merge();
                sheet.Range[sheet.Cells[i, 10], sheet.Cells[i, 12]].Merge();
                sheet.Range[sheet.Cells[i, 13], sheet.Cells[i, 14]].Merge();
                sheet.Range[sheet.Cells[i, 15], sheet.Cells[i, 16]].Merge();
                sheet.Range[sheet.Cells[i, 18], sheet.Cells[i, 21]].Merge();
                sheet.Range[sheet.Cells[i, 22], sheet.Cells[i, 24]].Merge();
                sheet.Range[sheet.Cells[i, 25], sheet.Cells[i, 26]].Merge();
            }
        }

        override protected void DrawBorders()
        {
            base.DrawBorders();
            // title
            sheet.Range[sheet.Cells[firstRow, 3], sheet.Cells[firstRow + 2, 26]].
                Borders.Weight = Excel.XlBorderWeight.xlMedium;
            // rows
            sheet.Range[sheet.Cells[firstRow + 3, 3], sheet.Cells[firstRow + 29, 26]].
                Borders.Weight = Excel.XlBorderWeight.xlMedium;
            sheet.Range[sheet.Cells[firstRow + 3, 3], sheet.Cells[firstRow + 29, 26]].
                Borders[Excel.XlBordersIndex.xlInsideHorizontal].Weight = Excel.XlBorderWeight.xlThin; // внутренние горизонтальные            
        }

        override protected void FillBlank()
        {
            base.FillBlank();

            // title default text
            sheet.Range[sheet.Cells[firstRow, 3], sheet.Cells[firstRow, 26]].
                Value2 = "Лист регистрации изменений";
            sheet.Range[sheet.Cells[firstRow + 1, 3], sheet.Cells[firstRow + 2, 4]].
                Value2 = "Изм.";
            sheet.Range[sheet.Cells[firstRow + 1, 5], sheet.Cells[firstRow + 1, 14]].
                Value2 = "Номера листов (страниц)";
            sheet.Range[sheet.Cells[firstRow + 2, 5], sheet.Cells[firstRow + 2, 7]].
                Value2 = "изменен- ных";
            sheet.Range[sheet.Cells[firstRow + 2, 8], sheet.Cells[firstRow + 2, 9]].
                Value2 = "заменен- ных";
            sheet.Range[sheet.Cells[firstRow + 2, 10], sheet.Cells[firstRow + 2, 12]].
                Value2 = "новых";
            sheet.Range[sheet.Cells[firstRow + 2, 13], sheet.Cells[firstRow + 2, 14]].
                Value2 = "аннулиро-ванных";
            sheet.Range[sheet.Cells[firstRow + 1, 15], sheet.Cells[firstRow + 2, 16]].
                Value2 = "Всего листов (страниц)    в докум.";
            sheet.Range[sheet.Cells[firstRow + 1, 17], sheet.Cells[firstRow + 2, 17]].
                Value2 = "№ докум.";
            sheet.Range[sheet.Cells[firstRow + 1, 18], sheet.Cells[firstRow + 2, 21]].
                Value2 = "Входящий №    сопроводительного    докум. и дата";
            sheet.Range[sheet.Cells[firstRow + 1, 22], sheet.Cells[firstRow + 2, 24]].
                Value2 = "Подп.";
            sheet.Range[sheet.Cells[firstRow + 1, 25], sheet.Cells[firstRow + 2, 26]].
                Value2 = "Дата";

            // font sizes
            // title
            sheet.Range[sheet.Cells[firstRow, 3], sheet.Cells[firstRow, 26]].Font.Size = 12;
            sheet.Range[sheet.Cells[firstRow + 1, 3], sheet.Cells[firstRow + 2, 4]].Font.Size = 12;
            sheet.Range[sheet.Cells[firstRow + 1, 5], sheet.Cells[firstRow + 1, 14]].Font.Size = 12;
            sheet.Range[sheet.Cells[firstRow + 2, 5], sheet.Cells[firstRow + 2, 12]].Font.Size = 11;
            sheet.Range[sheet.Cells[firstRow + 2, 13], sheet.Cells[firstRow + 2, 14]].Font.Size = 8;
            sheet.Range[sheet.Cells[firstRow + 1, 15], sheet.Cells[firstRow + 2, 16]].Font.Size = 10;
            sheet.Range[sheet.Cells[firstRow + 1, 17], sheet.Cells[firstRow + 2, 17]].Font.Size = 12;
            sheet.Range[sheet.Cells[firstRow + 1, 18], sheet.Cells[firstRow + 2, 21]].Font.Size = 9;
            sheet.Range[sheet.Cells[firstRow + 1, 22], sheet.Cells[firstRow + 2, 26]].Font.Size = 12;
            // rows
            sheet.Range[sheet.Cells[firstRow + 3, 3], sheet.Cells[firstRow + 29, 26]].Font.Size = 11;

            // text align
            // title
            sheet.Range[sheet.Cells[firstRow, 3], sheet.Cells[firstRow + 2, 26]].
                VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            sheet.Range[sheet.Cells[firstRow, 3], sheet.Cells[firstRow + 2, 26]].
                HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            // rows
            sheet.Range[sheet.Cells[firstRow + 3, 3], sheet.Cells[firstRow + 29, 16]].
                HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            sheet.Range[sheet.Cells[firstRow, 3], sheet.Cells[firstRow + 2, 26]].WrapText = true;
            sheet.Range[sheet.Cells[firstRow + 3, 3], sheet.Cells[firstRow + 29, 26]].
                ShrinkToFit = true;
        }
    }
}
